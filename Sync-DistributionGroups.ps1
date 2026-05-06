#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    Синхронизация групп рассылки AD/Exchange с источником подразделений (JSON + 1С).
.DESCRIPTION
    Создаёт/наполняет/удаляет группы рассылки на основе:
      - departments.json   — список актуальных подразделений;
      - ImportedDeps.csv   — иерархия (parent/child);
      - protected_groups.json — защита от удаления + аудит переименований;
      - Подразделения.xlsx — справочник 1С для сверки имён.
.NOTES
    Все операции выполняются на одном DC ($DcServer), read-after-write согласован.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$JsonFilePath            = 'C:\ProgramData\UniversalAutomation\departments.json',
    [string]$TargetOU                = 'OU=Группы рассылки,DC=transitcard,DC=ru',
    [string]$GroupDesc               = 'Группа рассылки подразделения_v1.0',
    [string]$MailDomain              = 'transitcard.ru',
    [string]$ExchangeServer          = 'srv-ex16-01.transitcard.ru',
    [string]$DcServer                = 'srv-dc3.transitcard.ru',
    [double]$FuzzyMatchThreshold     = 0.85,
    [string]$ProtectedGroupsJsonPath = 'D:\Scripts\Группы рассылки\protected_groups.json',
    [string]$ExcelPath               = 'D:\Scripts\verify1C\1c\Подразделения.xlsx',
    [string]$ReportCsvPath           = 'D:\Scripts\Группы рассылки\RenameReport.csv',
    [string]$LogPath                 = 'D:\Scripts\Группы рассылки\Logs',
    [int]$SamMaxLength               = 20
)

$ErrorActionPreference = 'Stop'

# ==========================================
# Кодировка консоли (только в интерактивном режиме)
# ==========================================
try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    [Console]::InputEncoding  = [System.Text.Encoding]::UTF8
} catch {
    Write-Warning "Не удалось установить кодировку консоли (возможно, скрипт запущен без консольного окна)."
}
$PSDefaultParameterValues['Get-Content:Encoding'] = 'UTF8'

# ==========================================
# Логирование
# ==========================================
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
$transcript = Join-Path $LogPath ("DG-Sync_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $transcript -Append | Out-Null

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $color = switch ($Level) {
        'ERROR' { 'Red' }; 'WARN' { 'Yellow' }; 'OK' { 'Green' }
        'STEP'  { 'Cyan' }; 'DEBUG' { 'DarkGray' }; default { 'Gray' }
    }
    Write-Host ("[{0:HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level, $Message) -ForegroundColor $color
}

# ==========================================
# Транслитерация (словарь — один раз на сессию)
# ==========================================
$script:CyrLatMap = @{}
$low = @{'а'='a';'б'='b';'в'='v';'г'='g';'д'='d';'е'='e';'ё'='e';'ж'='zh';'з'='z';'и'='i';'й'='y';'к'='k';'л'='l';'м'='m';'н'='n';'о'='o';'п'='p';'р'='r';'с'='s';'т'='t';'у'='u';'ф'='f';'х'='h';'ц'='ts';'ч'='ch';'ш'='sh';'щ'='sch';'ъ'='';'ы'='y';'ь'='';'э'='e';'ю'='yu';'я'='ya'}
$up  = @{'А'='A';'Б'='B';'В'='V';'Г'='G';'Д'='D';'Е'='E';'Ё'='E';'Ж'='Zh';'З'='Z';'И'='I';'Й'='Y';'К'='K';'Л'='L';'М'='M';'Н'='N';'О'='O';'П'='P';'Р'='R';'С'='S';'Т'='T';'У'='U';'Ф'='F';'Х'='H';'Ц'='Ts';'Ч'='Ch';'Ш'='Sh';'Щ'='Sch';'Ъ'='';'Ы'='Y';'Ь'='';'Э'='E';'Ю'='Yu';'Я'='Ya'}
foreach ($k in $low.Keys) { $script:CyrLatMap[[char]$k] = $low[$k] }
foreach ($k in $up.Keys)  { $script:CyrLatMap[[char]$k] = $up[$k] }

function Convert-CyrillicToLatin {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $sb = [System.Text.StringBuilder]::new($Text.Length)
    foreach ($ch in $Text.ToCharArray()) {
        if ($script:CyrLatMap.ContainsKey($ch)) { [void]$sb.Append($script:CyrLatMap[$ch]) }
        else { [void]$sb.Append($ch) }
    }
    return ($sb.ToString() -replace '[^a-zA-Z0-9\-_]', '')
}

function Get-Initials {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $sb = [System.Text.StringBuilder]::new()
    foreach ($word in ($Text -split '[\s\-]+')) {
        $clean = $word -replace '[^\p{L}0-9]', ''
        if ($clean.Length -gt 0) { [void]$sb.Append($clean[0]) }
    }
    return $sb.ToString()
}

function Normalize-Name {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return ($Text -replace '\s+', ' ').Trim().ToLowerInvariant()
}

# ==========================================
# Левенштейн с ранним отсечением
# ==========================================
function Get-StringSimilarity {
    param([string]$String1, [string]$String2)
    if ([string]::IsNullOrEmpty($String1) -and [string]::IsNullOrEmpty($String2)) { return 1.0 }
    if ([string]::IsNullOrEmpty($String1) -or  [string]::IsNullOrEmpty($String2)) { return 0.0 }
    $s1 = Normalize-Name $String1
    $s2 = Normalize-Name $String2
    $len1 = $s1.Length; $len2 = $s2.Length
    $maxLen = [Math]::Max($len1, $len2)
    if ($maxLen -eq 0) { return 1.0 }
    # Раннее отсечение по длине
    if ([Math]::Abs($len1 - $len2) / [double]$maxLen -gt 0.5) { return 0.0 }

    $prev = 0..$len2
    $curr = New-Object int[] ($len2 + 1)
    for ($i = 1; $i -le $len1; $i++) {
        $curr[0] = $i
        $c1 = $s1[$i - 1]
        for ($j = 1; $j -le $len2; $j++) {
            $cost = if ($c1 -eq $s2[$j - 1]) { 0 } else { 1 }
            $min  = $curr[$j - 1] + 1
            if ($prev[$j] + 1 -lt $min)         { $min = $prev[$j] + 1 }
            if ($prev[$j - 1] + $cost -lt $min) { $min = $prev[$j - 1] + $cost }
            $curr[$j] = $min
        }
        $tmp = $prev; $prev = $curr; $curr = $tmp
    }
    return 1 - ($prev[$len2] / $maxLen)
}

function Find-BestGroupMatch {
    param([string]$SearchName, [hashtable]$GroupsMap, [double]$Threshold = 0.85)
    if ([string]::IsNullOrWhiteSpace($SearchName)) { return $null }
    $key = Normalize-Name $SearchName
    if ($GroupsMap.ContainsKey($key)) { return $GroupsMap[$key] }
    $best = $null; $bestScore = 0.0
    foreach ($k in $GroupsMap.Keys) {
        $s = Get-StringSimilarity -String1 $key -String2 $k
        if ($s -ge $Threshold -and $s -gt $bestScore) {
            $bestScore = $s; $best = $GroupsMap[$k]
        }
    }
    return $best
}

function Get-UniqueSamAccountName {
    param([string]$BaseName, [hashtable]$Used, [int]$MaxLen = 20)
    $base = $BaseName
    if ($base.Length -gt $MaxLen) { $base = $base.Substring(0, $MaxLen) }
    $candidate = $base; $i = 1
    while ($Used.ContainsKey($candidate) -or
           (Get-ADGroup -Filter "SamAccountName -eq '$candidate'" -Server $DcServer -ErrorAction SilentlyContinue)) {
        $suffix = [string]$i
        $cut = $MaxLen - $suffix.Length
        if ($cut -lt 1) { throw "Не удалось подобрать SamAccountName для '$BaseName'" }
        $candidate = $base.Substring(0, [Math]::Min($base.Length, $cut)) + $suffix
        $i++
    }
    return $candidate
}

# ==========================================
# Подключение к Exchange
# ==========================================
$session = $null
try {
    Write-Log "Подключение к Exchange $ExchangeServer..." 'STEP'
    $session = New-PSSession -ConfigurationName Microsoft.Exchange `
                             -ConnectionUri ("http://{0}/PowerShell/" -f $ExchangeServer) `
                             -Authentication Kerberos
    Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
    Write-Log "Сессия Exchange открыта." 'OK'

    # ----------------------------------------
    # Загрузка защищённых групп
    # ----------------------------------------
    $protectedByEmail = @{}
    if (Test-Path $ProtectedGroupsJsonPath) {
        try {
            $protectedList = Get-Content -Path $ProtectedGroupsJsonPath -Raw | ConvertFrom-Json
            foreach ($entry in $protectedList) {
                if ($entry.email) {
                    $protectedByEmail[$entry.email.ToLowerInvariant().Trim()] = $entry
                }
            }
            Write-Log "Загружено $($protectedByEmail.Count) защищённых групп." 'OK'
        } catch {
            Write-Log "Ошибка чтения protected_groups.json: $($_.Exception.Message)" 'ERROR'
        }
    } else {
        Write-Log "protected_groups.json не найден — защита от удаления отключена." 'WARN'
    }

    # ==========================================
    # ШАГ 1. Создание групп из JSON
    # ==========================================
    Write-Log "Шаг 1: создание групп из JSON" 'STEP'
    if (-not (Test-Path $JsonFilePath)) { throw "JSON файл $JsonFilePath не найден." }

    $jsonContent = Get-Content -Path $JsonFilePath -Raw | ConvertFrom-Json
    $departments = @()
    if ($jsonContent -is [array] -and $jsonContent.Count -gt 0 -and $jsonContent[0] -is [string]) {
        $departments = $jsonContent
    } elseif ($jsonContent.department) {
        $departments = $jsonContent.department
    } else {
        $departments = $jsonContent | Select-Object -ExpandProperty department -ErrorAction SilentlyContinue
    }
    $departments = $departments |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { ($_ -replace '\s+', ' ').Trim() } |
        Select-Object -Unique

    # Один запрос: все управляемые группы в OU
    $existingGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                  -SearchBase $TargetOU -Server $DcServer `
                                  -Properties Name, SamAccountName, mail, DisplayName, extensionAttribute1, Member
    $groupsByName = @{}
    $usedSam      = @{}
    foreach ($g in $existingGroups) {
        $groupsByName[$g.Name.Trim()] = $g
        $usedSam[$g.SamAccountName]  = $true
    }

    foreach ($dept in $departments) {
        if ($groupsByName.ContainsKey($dept)) { continue }

        $groupName = $dept
        if ($groupName.Length -gt 64) {
            $groupName = $groupName.Substring(0, 64)
            Write-Log "Имя группы обрезано до 64 символов: '$groupName'" 'WARN'
        }

        $latin = Convert-CyrillicToLatin -Text (Get-Initials -Text $dept)
        $base  = ('DG_' + $latin).ToUpper()
        if ([string]::IsNullOrWhiteSpace($base)) { $base = 'DG' }

        try {
            $sam = Get-UniqueSamAccountName -BaseName $base -Used $usedSam -MaxLen $SamMaxLength
        } catch {
            Write-Log "Ошибка подбора SamAccountName для '$dept': $_" 'ERROR'
            continue
        }
        $usedSam[$sam] = $true

        if ($PSCmdlet.ShouldProcess($groupName, "New-ADGroup + Enable-DistributionGroup")) {
            try {
                New-ADGroup -Name $groupName -SamAccountName $sam `
                            -GroupCategory Distribution -GroupScope Universal `
                            -Path $TargetOU -Description $GroupDesc -DisplayName $groupName `
                            -Server $DcServer

                Enable-DistributionGroup -Identity $sam -Alias $sam `
                    -PrimarySmtpAddress ("{0}@{1}" -f $sam.ToLower(), $MailDomain) `
                    -DomainController $DcServer

                Set-DistributionGroup -Identity $sam `
                    -RequireSenderAuthenticationEnabled $false `
                    -DomainController $DcServer

                Write-Log "Создана группа: $groupName ($sam)" 'OK'
            } catch {
                Write-Log "Ошибка создания '$groupName': $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 1.1. ПОЛНАЯ ОЧИСТКА СОСТАВА ВСЕХ ГРУПП
    # ==========================================
    Write-Log "Шаг 1.1: полная очистка состава всех групп" 'STEP'
    $allGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" -SearchBase $TargetOU -Server $DcServer
    foreach ($g in $allGroups) {
        if ($PSCmdlet.ShouldProcess($g.Name, "Clear members")) {
            try {
                Set-ADGroup -Identity $g -Clear Member -Server $DcServer
            } catch {
                Write-Log "Ошибка очистки $($g.Name): $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # Перечитываем актуальный список (группы пусты)
    $managedGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                                 -SearchBase $TargetOU -Server $DcServer `
                                 -Properties Member, mail, DisplayName, extensionAttribute1
    $managedGroupsMap = @{}
    foreach ($g in $managedGroups) {
        $managedGroupsMap[(Normalize-Name $g.Name)] = $g
    }

    # ==========================================
    # Загрузка данных 1С (Excel) — иерархия, руководители, ID для аудита
    # ==========================================
    $excelDeptById  = @{}
    $excelHierarchy = New-Object System.Collections.Generic.List[object]
    $excelManagers  = New-Object System.Collections.Generic.List[object]
    if (Test-Path $ExcelPath) {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Log "Модуль ImportExcel не установлен. Установите: Install-Module ImportExcel -Scope CurrentUser" 'WARN'
        } else {
            try {
                Import-Module ImportExcel
                $excelData = Import-Excel -Path $ExcelPath -StartRow 2
                foreach ($row in $excelData) {
                    $childName  = if ($row.objFullName) { $row.objFullName.ToString().Trim() } else { $null }
                    $parentName = if ($row.objName)     { $row.objName.ToString().Trim() }     else { $null }
                    $managerEN  = if ($row.'Руководитель подразделения') { ($row.'Руководитель подразделения'.ToString().Trim() -replace '^0+', '') } else { $null }

                    if ($row.orgUnitId -and $childName) {
                        $excelDeptById[$row.orgUnitId.ToString().Trim()] = $childName
                    }
                    if ($childName -and $parentName -and ($childName -ne $parentName)) {
                        $excelHierarchy.Add([PSCustomObject]@{ Child = $childName; Parent = $parentName })
                    }
                    if ($childName -and $managerEN) {
                        $excelManagers.Add([PSCustomObject]@{ Department = $childName; EmployeeNumber = $managerEN })
                    }
                }
                Write-Log "1С: подразделений=$($excelDeptById.Count), иерархия=$($excelHierarchy.Count), руководителей=$($excelManagers.Count)" 'OK'
            } catch {
                Write-Log "Ошибка чтения Excel: $($_.Exception.Message)" 'ERROR'
            }
        }
    } else {
        Write-Log "Excel $ExcelPath не найден." 'WARN'
    }

    # ==========================================
    # ШАГ 2. Наполнение пользователями (только добавление)
    # ==========================================
    Write-Log "Шаг 2: добавление пользователей" 'STEP'

    $filterUsers = @"
Department -like '*'
  -and Enabled -eq 'True'
  -and EmployeeNumber -like '*'
  -and SamAccountName -notlike 'adm-*'
  -and SamAccountName -notlike 'vpn-*'
  -and SamAccountName -notlike '*test*'
  -and Name -notlike '*ADM*'
  -and Name -notlike '*VPN*'
  -and Name -notlike '*KZ*'
"@ -replace "`r`n", " " -replace "`n", " "

    $users = Get-ADUser -Filter $filterUsers `
                        -Properties Department, EmployeeNumber, Name `
                        -ResultPageSize 1000 -Server $DcServer

    $usersByDept   = @{}
    $usersByEmpNum = @{}
    foreach ($u in $users) {
        $key = Normalize-Name $u.Department
        if (-not $usersByDept.ContainsKey($key)) {
            $usersByDept[$key] = New-Object System.Collections.Generic.List[object]
        }
        $usersByDept[$key].Add($u)
        if ($u.EmployeeNumber) {
            $usersByEmpNum[$u.EmployeeNumber.ToString().Trim()] = $u
        }
    }

    foreach ($g in $managedGroups) {
        $key = Normalize-Name $g.Name
        $desired = if ($usersByDept.ContainsKey($key)) { $usersByDept[$key] } else { @() }
        if ($desired.Count -eq 0) { continue }
        $desiredDNs = @($desired | ForEach-Object { $_.DistinguishedName })
        $toAdd = $desiredDNs | Where-Object { $_ -and ($g.Member -notcontains $_) }
        if ($toAdd.Count -gt 0) {
            if ($PSCmdlet.ShouldProcess($g.Name, "Add $($toAdd.Count) users")) {
                try {
                    Add-ADGroupMember -Identity $g -Members $toAdd -Server $DcServer
                    Write-Log "+$($toAdd.Count) -> $($g.Name)" 'OK'
                } catch { Write-Log "Add fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
            }
        }
    }

    # ==========================================
    # ШАГ 2.1. Добавление руководителей подразделений из 1С
    # ==========================================
    Write-Log "Шаг 2.1: руководители подразделений" 'STEP'
    foreach ($m in $excelManagers) {
        if (-not $usersByEmpNum.ContainsKey($m.EmployeeNumber)) {
            Write-Log "Руководитель EmployeeNumber=$($m.EmployeeNumber) для '$($m.Department)' не найден в AD" 'WARN'
            continue
        }
        $manager = $usersByEmpNum[$m.EmployeeNumber]
        $group = Find-BestGroupMatch -SearchName $m.Department -GroupsMap $managedGroupsMap -Threshold $FuzzyMatchThreshold
        if (-not $group) {
            Write-Log "Группа подразделения '$($m.Department)' не найдена" 'WARN'
            continue
        }
        if ($PSCmdlet.ShouldProcess($group.Name, "Add manager $($manager.SamAccountName)")) {
            try {
                Add-ADGroupMember -Identity $group -Members $manager.DistinguishedName -Server $DcServer -ErrorAction Stop
                Write-Log "Руководитель $($manager.SamAccountName) -> $($group.Name)" 'OK'
            } catch {
                if ($_.Exception.Message -notmatch 'already a member') {
                    Write-Log "Add manager fail $($group.Name): $($_.Exception.Message)" 'ERROR'
                }
            }
        }
        if ($PSCmdlet.ShouldProcess($group.Name, "Set ManagedBy=$($manager.SamAccountName)")) {
            try {
                Set-ADGroup -Identity $group -ManagedBy $manager.DistinguishedName -Server $DcServer
                Write-Log "ManagedBy: $($group.Name) <- $($manager.SamAccountName)" 'OK'
            } catch {
                Write-Log "Set ManagedBy fail $($group.Name): $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 3. Иерархия из 1С (Excel)
    # ==========================================
    Write-Log "Шаг 3: иерархия (порог $FuzzyMatchThreshold)" 'STEP'
    foreach ($row in $excelHierarchy) {
        $parent = Find-BestGroupMatch -SearchName $row.Parent -GroupsMap $managedGroupsMap -Threshold $FuzzyMatchThreshold
        $child  = Find-BestGroupMatch -SearchName $row.Child  -GroupsMap $managedGroupsMap -Threshold $FuzzyMatchThreshold
        if (-not $parent) { Write-Log "Родитель '$($row.Parent)' не найден" 'WARN'; continue }
        if (-not $child)  { Write-Log "Потомок '$($row.Child)' не найден"  'WARN'; continue }
        try {
            if ($PSCmdlet.ShouldProcess("$($child.Name) -> $($parent.Name)", "Add child group")) {
                Add-ADGroupMember -Identity $parent -Members $child -Server $DcServer -ErrorAction Stop
                Write-Log "Иерархия: $($child.Name) -> $($parent.Name)" 'OK'
            }
        } catch {
            if ($_.Exception.Message -notmatch 'already a member') {
                Write-Log "Иерархия fail: $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    # ==========================================
    # ШАГ 3.1. Аудит переименований из 1С
    # ==========================================
    Write-Log "Шаг 3.1: аудит переименований" 'STEP'
    $report = New-Object System.Collections.Generic.List[object]
    foreach ($g in $managedGroups) {
        $email = if ($g.mail) { $g.mail.ToLowerInvariant().Trim() } else { $null }
        if (-not ($email -and $protectedByEmail.ContainsKey($email))) { continue }

        $currentDisplay = if ($g.DisplayName) { $g.DisplayName.Trim() } else { '' }
        $attrId = if ($g.extensionAttribute1) { $g.extensionAttribute1.Trim() } else { $null }
        $excelName = ''
        $status =
            if (-not $attrId) { 'Нет ID в extensionAttribute1' }
            elseif (-not $excelDeptById.ContainsKey($attrId)) { 'ID не найден в 1С' }
            else {
                $excelName = $excelDeptById[$attrId]
                if ($currentDisplay -ne $excelName) { 'Требуется переименование' }
                else { 'Имя актуально' }
            }
        $report.Add([PSCustomObject]@{
            Email        = $g.mail
            CurrentName  = $currentDisplay
            DepartmentId = $attrId
            NameIn1C     = $excelName
            Status       = $status
        })
    }
    if ($report.Count -gt 0) {
        $report | Export-Csv -Path $ReportCsvPath -NoTypeInformation -Encoding UTF8
        Write-Log "Отчёт: $ReportCsvPath ($($report.Count) записей)" 'OK'
    }

    # ==========================================
    # ШАГ 4. Удаление пустых групп (с учётом защиты)
    # ==========================================
    Write-Log "Шаг 4: удаление пустых групп" 'STEP'
    $finalGroups = Get-ADGroup -LDAPFilter "(description=$GroupDesc)" `
                               -SearchBase $TargetOU -Server $DcServer `
                               -Properties Member, mail
    foreach ($g in $finalGroups) {
        if ($g.Member -and $g.Member.Count -gt 0) { continue }
        $email = if ($g.mail) { $g.mail.ToLowerInvariant().Trim() } else { $null }
        if ($email -and $protectedByEmail.ContainsKey($email)) {
            Write-Log "Защищена JSON, пропуск: $($g.Name) ($email)" 'WARN'
            continue
        }
        if ($PSCmdlet.ShouldProcess($g.Name, "Disable + Remove")) {
            try {
                Disable-DistributionGroup -Identity $g.DistinguishedName -Confirm:$false -DomainController $DcServer
            } catch { Write-Log "Disable fail $($g.Name): $($_.Exception.Message)" 'WARN' }
            try {
                Remove-ADGroup -Identity $g -Confirm:$false -Server $DcServer
                Write-Log "Удалена пустая группа: $($g.Name)" 'OK'
            } catch { Write-Log "Remove fail $($g.Name): $($_.Exception.Message)" 'ERROR' }
        }
    }
}
catch {
    Write-Log "FATAL: $($_.Exception.Message)" 'ERROR'
    Write-Log $_.ScriptStackTrace 'ERROR'
    throw
}
finally {
    if ($session) {
        Remove-PSSession $session -ErrorAction SilentlyContinue
        Write-Log "Сессия Exchange закрыта." 'OK'
    }
    Write-Log "Готово (DC=$DcServer)." 'STEP'
    Stop-Transcript | Out-Null
}
