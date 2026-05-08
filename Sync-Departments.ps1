# Sync-Departments.ps1
# Экспорт уникальных названий подразделений активных пользователей AD в JSON-файл.
param(
    [string]$TargetServer = "transitcard.ru",
    [string]$OutputPath   = "C:\ProgramData\UniversalAutomation\departments.json"
)

Import-Module ActiveDirectory -ErrorAction Stop

try {
    # LDAP-фильтр:
    # - включённые пользователи: (!(userAccountControl:1.2.840.113556.1.4.803:=2))
    # - EmployeeNumber не пуст
    # - Department не пуст
    $ldapFilter = "(&(!(userAccountControl:1.2.840.113556.1.4.803:=2))(employeeNumber=*)(department=*))"

    $users = Get-ADUser -LDAPFilter $ldapFilter -Server $TargetServer `
                        -Properties Department, Name -ErrorAction Stop

    # Фильтрация на стороне PowerShell:
    # - исключаем имена с метками [ADM]/[VPN]/[KZ]
    # - проверяем, что Department содержит хотя бы один непробельный символ
    $filteredUsers = $users | Where-Object {
        $_.Name -notmatch '\[ADM\]' -and
        $_.Name -notmatch '\[VPN\]' -and
        $_.Name -notmatch '\[KZ\]' -and
        $_.Department -match '\S'
    }

    $departments = $filteredUsers |
                   Select-Object -ExpandProperty Department -Unique |
                   Sort-Object

    $dir = Split-Path $OutputPath -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    # `, @($departments)` форсирует массив даже при одном элементе,
    # иначе ConvertTo-Json вернёт скаляр и потребитель не распарсит.
    $json = , @($departments) | ConvertTo-Json
    Set-Content -Path $OutputPath -Value $json -Encoding UTF8

    Write-Host "Экспортировано уникальных отделов: $(@($departments).Count)" -ForegroundColor Green
} catch {
    Write-Error "Ошибка при получении отделов: $_"
    exit 1
}
