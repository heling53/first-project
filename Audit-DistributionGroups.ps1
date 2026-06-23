# ========== АБСОЛЮТНЫЕ ПУТИ К ФАЙЛАМ ==========
$TemplatePath = "D:\Scripts\DistributionGroups\template_message1.html"
$LogFile      = "D:\Scripts\DistributionGroups\task.log"
$ExcelPath    = "D:\Scripts\verify1C\1c\Подразделения.xlsx"
$ExcelSheet   = "Лист_1"

# CSV-файлы
$Csv_NonCorrADGroup        = "D:\Scripts\DistributionGroups\NonCorrADGroup.csv"
$Csv_NCMUpload             = "D:\Scripts\DistributionGroups\NCMUpload.csv"
$Csv_NonCorrUserMembership = "D:\Scripts\DistributionGroups\NonCorrUserMembership.csv"
$Csv_UserDepToAdd          = "D:\Scripts\DistributionGroups\UserDepToAdd.csv"
$Csv_UsersToExclude        = "D:\Scripts\DistributionGroups\UsersToExclude.csv"
$Csv_DGroupsToDel          = "D:\Scripts\DistributionGroups\DGroupsToDel.csv"
$Csv_DGroupsToRename       = "D:\Scripts\DistributionGroups\DGroupsToRename.csv"
$Csv_ImportedDeps          = "D:\Scripts\DistributionGroups\ImportedDeps.csv"

# Параметры почты
$MailServer    = "owa.transitcard.ru"
$MailPort      = 25
$MailSender    = "infosec-notifications@pprcard.ru"
$MailRecipient = @("itsecurity@pprcard.ru")
$MailSubject   = "[Группы рассылки] Менеджмент групп рассылки подразделений"

# Порог нечёткого сравнения (в процентах)
$FuzzyThreshold = 85

# ========== ПРОВЕРКА ПОВТОРНОГО ЗАПУСКА ==========
$mutex = New-Object System.Threading.Mutex($false, "Global\DistributionGroupsScript")
if (-not $mutex.WaitOne(0)) {
    $date = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LogFile -Value "$date [Warning] Скрипт уже выполняется. Выход."
    exit
}

# ========== ФУНКЦИЯ ЛОГИРОВАНИЯ ==========
function Out-Log {
    param([string]$Message)
    $date = Get-Date -Format "[yyyy-MM-dd HH:mm:ss]"
    Add-Content -Path $LogFile -Value "$date $Message"
}

# ========== ПРОВЕРКА НАЛИЧИЯ МОДУЛЕЙ ==========
$missingModules = @()
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    $missingModules += "ActiveDirectory"
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    $missingModules += "ImportExcel"
}
if ($missingModules.Count -gt 0) {
    $errMsg = "Отсутствуют требуемые модули: $($missingModules -join ', ')"
    Out-Log "[Error] $errMsg"
    Write-Error $errMsg
    exit 1
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    Out-Log "[Information] Модуль ActiveDirectory загружен."
} catch {
    Out-Log "[Error] Не удалось загрузить модуль ActiveDirectory: $($_.Exception.Message)"
    exit 1
}

try {
    Import-Module ImportExcel -ErrorAction Stop
    Out-Log "[Information] Модуль ImportExcel загружен."
} catch {
    Out-Log "[Error] Не удалось загрузить модуль ImportExcel: $($_.Exception.Message)"
    exit 1
}

# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (СХОДСТВО СТРОК) ==========
Function Get-StringSimilarity {
    Param([string]$str1, [string]$str2)

    if ($str1 -eq $str2) { return 100 }
    if ([string]::IsNullOrWhiteSpace($str1) -or [string]::IsNullOrWhiteSpace($str2)) { return 0 }

    $s1 = $str1.ToLower().Trim()
    $s2 = $str2.ToLower().Trim()

    $len1 = $s1.Length
    $len2 = $s2.Length
    $matrix = New-Object 'int[,]' ($len1 + 1), ($len2 + 1)

    for ($i = 0; $i -le $len1; $i++) { $matrix[$i, 0] = $i }
    for ($j = 0; $j -le $len2; $j++) { $matrix[0, $j] = $j }

    for ($i = 1; $i -le $len1; $i++) {
        for ($j = 1; $j -le $len2; $j++) {
            $cost = if ($s1[$i - 1] -eq $s2[$j - 1]) { 0 } else { 1 }

            $deletion    = $matrix[($i - 1), $j] + 1
            $insertion   = $matrix[$i, ($j - 1)] + 1
            $substitution = $matrix[($i - 1), ($j - 1)] + $cost

            $minVal = [math]::Min($deletion, $insertion)
            $matrix[$i, $j] = [math]::Min($minVal, $substitution)
        }
    }

    $dist = $matrix[$len1, $len2]
    $maxL = [math]::Max($len1, $len2)
    return [math]::Round(((1 - ($dist / $maxL)) * 100), 2)
}

Function Find-BestGroupMatch {
    Param([string]$SearchName, [hashtable]$Map)

    if ([string]::IsNullOrWhiteSpace($SearchName)) { return $null }
    if ($Map.ContainsKey($SearchName)) { return $Map[$SearchName] }

    $bestMatch = $null
    $highestScore = 0

    foreach ($key in $Map.Keys) {
        $score = Get-StringSimilarity -str1 $SearchName -str2 $key
        if ($score -ge $FuzzyThreshold -and $score -gt $highestScore) {
            $highestScore = $score
            $bestMatch = $Map[$key]
        }
    }
    return $bestMatch
}

# ========== ПРОВЕРКА ЭКВИВАЛЕНТНОСТИ ИМЁН ГРУПП (УЧЁТ ОБРЕЗАНИЯ AD) ==========
function Test-GroupNameEquivalence {
    param(
        [string]$Name1,
        [string]$Name2
    )
    if ($Name1 -eq $Name2) { return $true }

    $len1 = $Name1.Length
    $len2 = $Name2.Length
    if ($len1 -gt 0 -and $len2 -gt 0) {
        $minLen = [Math]::Min($len1, $len2)
        $prefix = $Name1.Substring(0, $minLen)
        if ($Name2.StartsWith($prefix) -and [Math]::Abs($len1 - $len2) -le 5) {
            return $true
        }
    }

    $similarity = Get-StringSimilarity -str1 $Name1 -str2 $Name2
    return ($similarity -ge 95)
}

# ========== УНИВЕРСАЛЬНАЯ ФУНКЦИЯ ОТПРАВКИ ПИСЕМ ==========
function Send-Notification {
    param(
        [string]$MessageText = "",
        [string]$TableHtml = "",
        [array]$Attachments = @()
    )
    $MailTemplate = Get-Content -Path $TemplatePath -Raw -ErrorAction SilentlyContinue
    if (-not $MailTemplate) {
        $MailTemplate = "<html><body>{Message}<br>{Table}</body></html>"
    }
    $MailTemplate = $MailTemplate.Replace('{Message}', $MessageText).Replace('{Table}', $TableHtml)

    $MailMessage = @{
        From        = $MailSender
        To          = $MailRecipient
        Subject     = $MailSubject
        Body        = $MailTemplate
        SmtpServer  = $MailServer
        Port        = $MailPort
        Encoding    = "UTF8"
    }

    if ($Attachments.Count -gt 0) {
        $MailMessage.Attachments = $Attachments
    }

    try {
        Send-MailMessage @MailMessage -BodyAsHtml
        Out-Log "[Information] Notification email sent successfully."
    } catch {
        Out-Log "[Error] Failed to send notification email: $($_.Exception.Message)"
    }
}

# ========== ФУНКЦИЯ ИМПОРТА ДАННЫХ ИЗ EXCEL (ImportExcel) С ОЧИСТКОЙ ==========
function Import-DepartmentsFromExcel {
    param(
        [string]$FilePath,
        [string]$SheetName
    )
    try {
        if (-not (Test-Path $FilePath)) {
            throw "Файл '$FilePath' не найден."
        }

        $rawData = Import-Excel -Path $FilePath -WorksheetName $SheetName -ErrorAction Stop

        if ($rawData.Count -eq 0) {
            Out-Log "[Warning] Excel sheet '$SheetName' is empty or contains no data rows."
            return @()
        }

        $result = [System.Collections.Generic.List[PSObject]]::new()
        foreach ($row in $rawData) {
            $id        = $row.'Уникальный идентификатор подразделения'
            $fullName  = $row.'Полное наименование пдоразделения'
            $headName  = $row.'Наименование подразделения (вершины)'
            $managerEN = $row.'Руководитель подразделения'

            if (-not $fullName) {
                $fullName = $row.PSObject.Properties | Where-Object { $_.Name -like '*полное*' } | Select-Object -First 1 -ExpandProperty Value
            }
            if (-not $id) {
                $id = $row.PSObject.Properties | Where-Object { $_.Name -like '*уникальный*' -or $_.Name -like '*идентификатор*' } | Select-Object -First 1 -ExpandProperty Value
            }
            if (-not $headName) {
                $headName = $row.PSObject.Properties | Where-Object { $_.Name -like '*вершины*' } | Select-Object -First 1 -ExpandProperty Value
            }
            if (-not $managerEN) {
                $managerEN = $row.PSObject.Properties | Where-Object { $_.Name -like '*руководитель*' } | Select-Object -First 1 -ExpandProperty Value
            }

            $fullName = if ($fullName) { $fullName.ToString().Trim() } else { '' }
            if ([string]::IsNullOrWhiteSpace($fullName)) { continue }

            # Срезаем ведущие нули у EmployeeNumber руководителя (Excel-формат "0000003966").
            # Исключение: руководитель с табельным номером "0000-00013" — единственный, у кого
            # номер содержит дефис и ведущие нули значимы (в AD он хранится как "0000-00013").
            # Для него оставляем значение как есть, иначе оно превратится в "-00013" и не совпадёт с AD.
            $managerEN = if ($managerEN) {
                $en = $managerEN.ToString().Trim()
                if ($en -eq '0000-00013') { $en } else { $en -replace '^0+', '' }
            } else { '' }

            $obj = [PSCustomObject]@{
                'Уникальный идентификатор подразделения' = if ($id) { $id.ToString().Trim() } else { '' }
                'Полное наименование пдоразделения'      = $fullName
                'Наименование подразделения (вершины)'   = if ($headName) { $headName.ToString().Trim() } else { '' }
                'Руководитель подразделения'             = $managerEN
            }
            $result.Add($obj)
        }

        # Фильтрация служебного мусора и дубликатов
        $filtered = $result | Where-Object {
            $_.'Полное наименование пдоразделения' -notmatch 'objName|полное наименование|наименование|^obj'
        }
        $unique = $filtered | Group-Object -Property 'Полное наименование пдоразделения' | ForEach-Object { $_.Group[0] }

        Out-Log "[Information] Imported $($unique.Count) unique departments from Excel (total raw rows: $($result.Count))."
        return $unique
    } catch {
        Out-Log "[Error] Excel import failed: $($_.Exception.Message)"
        throw
    }
}

# ========== ФУНКЦИЯ ПОЛУЧЕНИЯ ДАННЫХ ИЗ AD ==========
function Get-ADData {
    try {
        $OUs = @("OU=Users.Contractor,DC=transitcard,DC=ru", "OU=Users.Corp,DC=transitcard,DC=ru")

        $AllADUsers = [System.Collections.Generic.List[object]]::new()
        foreach ($ou in $OUs) {
            $filter = "Enabled -eq 'True' -and EmployeeNumber -like '*' -and Name -notlike '*[ADM]*' -and Name -notlike '*[VPN]*' -and Name -notlike '*[KZ]*'"
            $users = Get-ADUser -SearchBase $ou -Filter $filter -Properties EmployeeNumber, Department, Name, SamAccountName, Enabled, DistinguishedName
            foreach ($u in $users) { $AllADUsers.Add($u) }
        }
        Out-Log "[Information] Loaded $($AllADUsers.Count) AD users."

        $AllADGroups = Get-ADGroup -Filter "Description -eq 'Группа рассылки подразделения_v1.0'" -Properties DisplayName, Description, SamAccountName, CN, DistinguishedName, Members, MemberOf, mail
        Out-Log "[Information] Loaded $($AllADGroups.Count) target AD groups."

        $GroupByDisplayName = @{}
        $GroupByCN          = @{}
        $GroupByDistName    = @{}
        $GroupBySamName     = @{}
        foreach ($g in $AllADGroups) {
            if ($g.DisplayName)       { $GroupByDisplayName[$g.DisplayName] = $g }
            if ($g.CN)                { $GroupByCN[$g.CN] = $g }
            if ($g.DistinguishedName) { $GroupByDistName[$g.DistinguishedName] = $g }
            if ($g.SamAccountName)    { $GroupBySamName[$g.SamAccountName] = $g }
        }

        $UserBySam = @{}
        $UserByDN  = @{}
        $UserByEmpNum = @{}
        foreach ($u in $AllADUsers) {
            $UserBySam[$u.SamAccountName] = $u
            $UserByDN[$u.DistinguishedName] = $u
            if ($u.EmployeeNumber) {
                $UserByEmpNum[$u.EmployeeNumber.ToString().Trim()] = $u
            }
        }

        return @{
            Users                = $AllADUsers
            Groups               = $AllADGroups
            GroupByDisplayName   = $GroupByDisplayName
            GroupByCN            = $GroupByCN
            GroupByDistName      = $GroupByDistName
            GroupBySamName       = $GroupBySamName
            UserBySam            = $UserBySam
            UserByDN             = $UserByDN
            UserByEmpNum         = $UserByEmpNum
        }
    } catch {
        Out-Log "[Error] AD data retrieval failed: $($_.Exception.Message)"
        throw
    }
}

# ========== ФУНКЦИЯ ПОИСКА НЕКОРРЕКТНЫХ ГРУПП ==========
function Find-NonCorrADGroup {
    param(
        $AllNotes,
        $GroupByDisplayName
    )
    $NonCorrADGroup = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Department in $AllNotes) {
        $depName = $Department."Полное наименование пдоразделения"
        if ($depName -match '^obj') { continue }

        $foundGroup = Find-BestGroupMatch -SearchName $depName -Map $GroupByDisplayName
        if (-not $foundGroup) {
            $adGroup = Get-ADGroup -Filter "DisplayName -eq '$depName'" -Properties Description -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($adGroup) {
                $NonCorrADGroup.Add([PSCustomObject]@{
                    DisplayName = $depName
                    Description = "Группа рассылки подразделения_v1.0"
                    Reason      = "Некорректное поле Description"
                })
            } else {
                $NonCorrADGroup.Add([PSCustomObject]@{
                    DisplayName = $depName
                    Description = "Группа рассылки подразделения_v1.0"
                    Reason      = "Группа с таким DN не найдена (Порог $FuzzyThreshold%)"
                })
            }
        }
    }
    $NonCorrADGroup | Export-Csv -Path $Csv_NonCorrADGroup -Encoding UTF8 -NoTypeInformation
    return $NonCorrADGroup
}

# ========== ФУНКЦИЯ ПРОВЕРКИ ИЕРАРХИИ (С УЧЁТОМ ОБРЕЗАНИЯ ИМЁН) ==========
function Test-GroupHierarchy {
    param(
        $AllNotes,
        $GroupByDisplayName,
        $GroupByDistName
    )
    $NCMUpload = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Department in $AllNotes) {
        $CurrDep = $Department.'Полное наименование пдоразделения'
        $HeadDep = $Department.'Наименование подразделения (вершины)'
        if ($CurrDep -eq "Руководство") { continue }
        if ([string]::IsNullOrWhiteSpace($HeadDep)) { continue }

        $targetGroup = Find-BestGroupMatch -SearchName $CurrDep -Map $GroupByDisplayName
        $parentGroup = Find-BestGroupMatch -SearchName $HeadDep -Map $GroupByDisplayName
        if (-not $targetGroup -or -not $parentGroup) { continue }

        $isMember = $false
        foreach ($memberOfDN in $targetGroup.MemberOf) {
            $parentCandidate = $GroupByDistName[$memberOfDN]
            if ($parentCandidate -and $parentCandidate.Description -eq "Группа рассылки подразделения_v1.0") {
                if (Test-GroupNameEquivalence -Name1 $parentCandidate.DisplayName -Name2 $HeadDep) {
                    $isMember = $true
                    break
                }
            }
        }

        $groupsToExclude = @()
        if ($isMember) {
            foreach ($dn in $targetGroup.MemberOf) {
                $parent = $GroupByDistName[$dn]
                if ($parent -and $parent.Description -eq "Группа рассылки подразделения_v1.0") {
                    if (-not (Test-GroupNameEquivalence -Name1 $parent.DisplayName -Name2 $HeadDep)) {
                        $groupsToExclude += $parent.DisplayName
                    }
                }
            }
            if ($groupsToExclude.Count) {
                $NCMUpload.Add([PSCustomObject]@{
                    DGroup     = $CurrDep
                    WhereToAdd = $HeadDep
                    DeleteFrom = $groupsToExclude -join ','
                })
            }
        } else {
            $NCMUpload.Add([PSCustomObject]@{
                DGroup     = $CurrDep
                WhereToAdd = $HeadDep
                DeleteFrom = ''
            })
        }
    }
    $NCMUpload | Export-Csv -Path $Csv_NCMUpload -Encoding UTF8 -NoTypeInformation
    return $NCMUpload
}

# ========== ФУНКЦИЯ ПОИСКА НЕКОРРЕКТНЫХ ЧЛЕНОВ (ПОЛЬЗОВАТЕЛЕЙ) ==========
function Find-NonCorrUserMembership {
    param(
        $AllNotes,
        $GroupByDisplayName,
        $GroupByDistName,
        $UserByDN
    )

    # Карта: нормализованное имя группы -> EmployeeNumber руководителя (без ведущих нулей).
    # Один человек может руководить несколькими подразделениями — его DN будет в нескольких группах,
    # и в каждой из них он не должен попадать в NonCorrUserMembership, даже если AD.Department не совпадает.
    $managersByDept = @{}
    foreach ($d in $AllNotes) {
        $dn = $d.'Полное наименование пдоразделения'
        $en = $d.'Руководитель подразделения'
        if ($dn -and $en) {
            $managersByDept[$dn] = $en
        }
    }

    $NonCorrUserMembership = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Department in $AllNotes) {
        $depName = $Department.'Полное наименование пдоразделения'
        $group = Find-BestGroupMatch -SearchName $depName -Map $GroupByDisplayName
        if (-not $group) { continue }

        $managerEN = $managersByDept[$depName]

        foreach ($memberDN in $group.Members) {
            $memberGroup = $GroupByDistName[$memberDN]
            if ($memberGroup) { continue }
            $user = $UserByDN[$memberDN]
            if (-not $user) { continue }

            # Если пользователь — руководитель этого подразделения (по EmployeeNumber из 1С),
            # пропускаем, даже если его AD.Department отличается (кейс: один человек руководит двумя подразделениями).
            if ($managerEN -and $user.EmployeeNumber `
                -and ($user.EmployeeNumber.ToString().Trim() -eq $managerEN)) {
                continue
            }

            if ((Get-StringSimilarity -str1 $user.Department -str2 $depName) -lt $FuzzyThreshold) {
                $NonCorrUserMembership.Add([PSCustomObject]@{
                    SamAccountName  = $user.SamAccountName
                    DeleteFromGroup = $depName
                })
            }
        }
    }
    $uniqueResult = $NonCorrUserMembership | Select-Object -Unique SamAccountName, DeleteFromGroup
    $uniqueResult | Export-Csv -Path $Csv_NonCorrUserMembership -Encoding UTF8 -NoTypeInformation
    return $uniqueResult
}

# ========== ФУНКЦИЯ ПОИСКА ПОЛЬЗОВАТЕЛЕЙ ДЛЯ ДОБАВЛЕНИЯ ==========
function Find-UsersToAdd {
    param(
        $AllADUsers,
        $GroupByCN
    )
    $UserDepToAdd = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($User in $AllADUsers) {
        $group = Find-BestGroupMatch -SearchName $User.Department -Map $GroupByCN
        if (-not $group) { continue }
        $isMember = $group.Members -contains $User.DistinguishedName
        if (-not $isMember) {
            $UserDepToAdd.Add([PSCustomObject]@{
                SamAccountName = $User.SamAccountName
                GroupToAdd     = $group.CN
            })
        }
    }
    $UserDepToAdd | Export-Csv -Path $Csv_UserDepToAdd -Encoding UTF8 -NoTypeInformation
    return $UserDepToAdd
}

# ========== ФУНКЦИЯ ПОИСКА НЕАКТИВНЫХ ПОЛЬЗОВАТЕЛЕЙ ДЛЯ УДАЛЕНИЯ ==========
function Find-UsersToExclude {
    param(
        $DGroups,
        $UserByDN
    )
    $UsersToExclude = [System.Collections.Generic.List[PSObject]]::new()
    foreach ($Group in $DGroups) {
        foreach ($memberDN in $Group.Members) {
            $user = $UserByDN[$memberDN]
            if (-not $user) { continue }
            if ($user.Enabled -ne $true) {
                $UsersToExclude.Add([PSCustomObject]@{
                    SamAccountName  = $user.SamAccountName
                    DeleteFromGroup = $Group.DisplayName
                })
            }
        }
    }
    $UsersToExclude | Export-Csv -Path $Csv_UsersToExclude -Encoding UTF8 -NoTypeInformation
    return $UsersToExclude
}

# ========== ФУНКЦИЯ ПОИСКА УСТАРЕВШИХ ГРУПП ДЛЯ УДАЛЕНИЯ / ПЕРЕИМЕНОВАНИЯ ==========
function Find-ObsoleteGroups {
    param(
        $DGroups,
        $AllNotes,
        $ImportedDeps
    )
    $DGroupsToDel = [System.Collections.Generic.List[PSObject]]::new()
    $DGroupsToRename = [System.Collections.Generic.List[PSObject]]::new()

    $currentDepNames = $AllNotes.'Полное наименование пдоразделения'
    $currentDepByDisplayName = @{}
    foreach ($dep in $AllNotes) {
        $currentDepByDisplayName[$dep.'Полное наименование пдоразделения'] = $dep
    }

    foreach ($Group in $DGroups) {
        $existsNow = $false
        foreach ($name in $currentDepNames) {
            if ((Get-StringSimilarity -str1 $Group.CN -str2 $name) -ge $FuzzyThreshold) {
                $existsNow = $true
                break
            }
        }

        if (-not $existsNow) {
            $foundCurrent = $false
            foreach ($key in $currentDepByDisplayName.Keys) {
                if ((Get-StringSimilarity -str1 $Group.DisplayName -str2 $key) -ge $FuzzyThreshold) {
                    $foundCurrent = $true
                    break
                }
            }
            if ($foundCurrent) {
                continue
            }

            $oldDep = $ImportedDeps | Where-Object { (Get-StringSimilarity -str1 $_.'Полное наименование пдоразделения' -str2 $Group.CN) -ge $FuzzyThreshold } | Select-Object -First 1
            if ($oldDep) {
                $newDep = $AllNotes | Where-Object { $_.'Уникальный идентификатор подразделения' -eq $oldDep.'Уникальный идентификатор подразделения' } | Select-Object -First 1
                if ($newDep -and ($newDep.'Полное наименование пдоразделения' -ne $oldDep.'Полное наименование пдоразделения')) {
                    $DGroupsToRename.Add([PSCustomObject]@{
                        Mail        = $Group.mail
                        OldDepName  = $oldDep.'Полное наименование пдоразделения'
                        NewDepName  = $newDep.'Полное наименование пдоразделения'
                    })
                } elseif (-not $newDep) {
                    $DGroupsToDel.Add([PSCustomObject]@{
                        Mail = $Group.mail
                        DN   = $Group.DisplayName
                    })
                }
            }
        }
    }
    $DGroupsToDel    | Export-Csv -Path $Csv_DGroupsToDel -Encoding UTF8 -NoTypeInformation
    $DGroupsToRename | Export-Csv -Path $Csv_DGroupsToRename -Encoding UTF8 -NoTypeInformation
    return @{
        ToDel    = $DGroupsToDel
        ToRename = $DGroupsToRename
    }
}

# ========== ФУНКЦИЯ ФОРМИРОВАНИЯ HTML-ТАБЛИЦЫ СВОДКИ ==========
# Null-безопасный подсчёт записей.
# Нужен потому, что функции возвращают пустой набор как $null (пустой List при выводе
# разворачивается в 0 объектов), а @($null).Count = 1, а не 0. Measure-Object даёт
# корректно: $null/пусто -> 0, скаляр -> 1, массив -> N.
function Get-RecordCount {
    param($Data)
    return ($Data | Measure-Object).Count
}

function Get-SummaryHtmlTable {
    param(
        $NonCorrADGroup,
        $NCMUpload,
        $NonCorrUserMembership,
        $UserDepToAdd,
        $UsersToExclude,
        $DGroupsToDel,
        $DGroupsToRename
    )
    $rows = @(
        @{ Файл = "NonCorrADGroup.csv";         Описание = "Некорректно добавленные или несуществующие группы рассылки подразделений";          Записей = (Get-RecordCount $NonCorrADGroup) },
        @{ Файл = "NCMUpload.csv";              Описание = "Записи о некорректном вхождении группы в вышестоящий список рассылки";                Записей = (Get-RecordCount $NCMUpload) },
        @{ Файл = "NonCorrUserMembership.csv";  Описание = "Записи о некорректном вхождении пользователя в группу рассылки";                      Записей = (Get-RecordCount $NonCorrUserMembership) },
        @{ Файл = "UserDepToAdd.csv";           Описание = "Записи для включения пользователей в группы рассылки";                                Записей = (Get-RecordCount $UserDepToAdd) },
        @{ Файл = "UsersToExclude.csv";         Описание = "Записи о деактивированных пользователях, подлежащих исключению из группы";            Записей = (Get-RecordCount $UsersToExclude) },
        @{ Файл = "DGroupsToDel.csv";           Описание = "Список групп рассылки, подлежащих удалению";                                          Записей = (Get-RecordCount $DGroupsToDel) },
        @{ Файл = "DGroupsToRename.csv";        Описание = "Группы, требующие переименования";                                                    Записей = (Get-RecordCount $DGroupsToRename) }
    )
    $html = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
    $html += "<tr><th>Файл</th><th>Описание</th><th>Количество записей</th></tr>"
    foreach ($row in $rows) {
        $html += "<tr><td>$($row.Файл)</td><td>$($row.Описание)</td><td style='text-align:center'>$($row.Записей)</td></tr>"
    }
    $html += "</table>"
    $html += "<br><b>Внимание!</b> Если добавлено новое подразделение и группа рассылки для него ранее не создавалась, необходимо придумать адрес группы рассылки и создать группу."
    return $html
}

# ========== ОСНОВНОЙ БЛОК СКРИПТА ==========
try {
    Out-Log "=== Начало выполнения ==="

    # 1. Импорт данных из Excel
    $AllNotes = Import-DepartmentsFromExcel -FilePath $ExcelPath -SheetName $ExcelSheet

    # 2. Загрузка данных из AD
    $ADData = Get-ADData
    $AllADUsers          = $ADData.Users
    $AllADGroups         = $ADData.Groups
    $GroupByDisplayName  = $ADData.GroupByDisplayName
    $GroupByCN           = $ADData.GroupByCN
    $GroupByDistName     = $ADData.GroupByDistName
    $UserByDN            = $ADData.UserByDN

    # 3. Некорректные группы рассылки
    $NonCorrADGroup = Find-NonCorrADGroup -AllNotes $AllNotes -GroupByDisplayName $GroupByDisplayName

    # 4. Проверка иерархии (с учётом обрезки имён)
    $NCMUpload = Test-GroupHierarchy -AllNotes $AllNotes -GroupByDisplayName $GroupByDisplayName -GroupByDistName $GroupByDistName

    # 5. Некорректные члены (пользователи) — с учётом руководителей нескольких подразделений
    $NonCorrUserMembership = Find-NonCorrUserMembership -AllNotes $AllNotes -GroupByDisplayName $GroupByDisplayName -GroupByDistName $GroupByDistName -UserByDN $UserByDN

    # 6. Пользователи для добавления
    $UserDepToAdd = Find-UsersToAdd -AllADUsers $AllADUsers -GroupByCN $GroupByCN

    # 7. Неактивные пользователи для исключения
    $DGroups = $AllADGroups
    $UsersToExclude = Find-UsersToExclude -DGroups $DGroups -UserByDN $UserByDN

    # 8. Устаревшие группы для удаления/переименования
    $ImportedDeps = if (Test-Path $Csv_ImportedDeps) { Import-Csv $Csv_ImportedDeps -Encoding UTF8 } else { @() }
    $ObsoleteGroups = Find-ObsoleteGroups -DGroups $DGroups -AllNotes $AllNotes -ImportedDeps $ImportedDeps
    $DGroupsToDel    = $ObsoleteGroups.ToDel
    $DGroupsToRename = $ObsoleteGroups.ToRename

    # 9. Сохранение текущего состояния подразделений
    $AllNotes | Export-Csv -Path $Csv_ImportedDeps -Encoding UTF8 -NoTypeInformation

    # 10. Формирование сводной HTML-таблицы
    $SummaryTableHtml = Get-SummaryHtmlTable -NonCorrADGroup $NonCorrADGroup -NCMUpload $NCMUpload -NonCorrUserMembership $NonCorrUserMembership -UserDepToAdd $UserDepToAdd -UsersToExclude $UsersToExclude -DGroupsToDel $DGroupsToDel -DGroupsToRename $DGroupsToRename

    # 11. Подготовка вложений (только файлы, в которых есть данные)
    $Attachments = @()
    if ((Get-RecordCount $NonCorrADGroup) -gt 0)        { $Attachments += $Csv_NonCorrADGroup }
    if ((Get-RecordCount $NCMUpload) -gt 0)             { $Attachments += $Csv_NCMUpload }
    if ((Get-RecordCount $NonCorrUserMembership) -gt 0) { $Attachments += $Csv_NonCorrUserMembership }
    if ((Get-RecordCount $UserDepToAdd) -gt 0)          { $Attachments += $Csv_UserDepToAdd }
    if ((Get-RecordCount $UsersToExclude) -gt 0)        { $Attachments += $Csv_UsersToExclude }
    if ((Get-RecordCount $DGroupsToDel) -gt 0)          { $Attachments += $Csv_DGroupsToDel }
    if ((Get-RecordCount $DGroupsToRename) -gt 0)       { $Attachments += $Csv_DGroupsToRename }

    # 12. Отправка итогового письма
    Send-Notification -TableHtml $SummaryTableHtml -Attachments $Attachments

    Out-Log "=== Успешное завершение ==="
} catch {
    Out-Log "[Critical] Script failed: $($_.Exception.Message)"
    exit 1
} finally {
    if ($mutex) {
        $mutex.ReleaseMutex()
    }
}
