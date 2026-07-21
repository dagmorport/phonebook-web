<#
.SYNOPSIS
    Export employees from Active Directory to JS file.
    Departments are exported AS-IS from AD.
#>

try {

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.DirectoryServices
Add-Type -AssemblyName System.DirectoryServices.AccountManagement

$dcServer   = 'srv-dc-002.e5dag.ru'

# ============================================================
# AUTHENTICATION — только Domain Admins
# ============================================================
Write-Host ""
Write-Host "========================================"
Write-Host "  Export Phonebook from AD"
Write-Host "========================================"
Write-Host ""
Write-Host "  Требуется аутентификация для экспорта из AD"
Write-Host ""

$cred = Get-Credential -Message "Введите учётные данные доменного администратора"
if ($null -eq $cred) {
    Write-Host "  ОШИБКА: Отмена ввода" -ForegroundColor Red
    exit 1
}

$login = $cred.UserName
$pass  = $cred.GetNetworkCredential().Password

# Извлекаем чистый логин (без домена, если пользователь ввёл DOMAIN\user или user@domain)
$loginClean = $login
if ($login -match '\\') { $loginClean = $login.Split('\')[1] }
if ($login -match '@')  { $loginClean = $login.Split('@')[0] }

Write-Host "  Проверка: $loginClean ..." -ForegroundColor Gray
Write-Host "  Контроллер домена: $dcServer" -ForegroundColor Gray

try {
    $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext(
        [System.DirectoryServices.AccountManagement.ContextType]::Domain, $dcServer
    )
    Write-Host "  Подключение к $dcServer выполнено" -ForegroundColor Gray
    $valid = $pc.ValidateCredentials($loginClean, $pass)
    if (-not $valid) {
        Write-Host "  ОШИБКА: Неверный логин или пароль" -ForegroundColor Red
        $pc.Dispose()
        exit 1
    }

    $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::FindByIdentity($pc, $loginClean)
    if ($null -eq $user) {
        Write-Host "  ОШИБКА: Пользователь не найден" -ForegroundColor Red
        $pc.Dispose()
        exit 1
    }

    # Проверка группы Domain Admins через whoami (локально, без доп. подключения к DC)
    $user.Dispose()
    $pc.Dispose()

    $whoamiOutput = whoami /groups 2>$null
    $isAdmin = $false
    if ($whoamiOutput -match 'Domain Admins' -or $whoamiOutput -match 'Администраторы домена') {
        $isAdmin = $true
    }

    if (-not $isAdmin) {
        Write-Host "  ОШИБКА: Нет прав. Требуется членство в группе Domain Admins" -ForegroundColor Red
        exit 1
    }

    Write-Host "  Аутентификация успешна: $loginClean (Domain Admin)" -ForegroundColor Green
    Write-Host ""

} catch {
    Write-Host "  ОШИБКА аутентификации: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

$serverPath = '\\srv-nas-001\DCSWHttpRepository\swuploads\Phonebook'
$basePath   = $serverPath
$outputFile = Join-Path $basePath 'employees.js'
$configFile = Join-Path $basePath 'dept-filter.json'

Write-Host "========================================"
Write-Host "========================================"
Write-Host ""

# ============================================================
# STEP 1: Query AD
# ============================================================
Write-Host "[1/5] Connecting to AD ($dcServer)..."

$searcher = New-Object System.DirectoryServices.DirectorySearcher
$searcher.Filter = "(&(objectCategory=person)(objectClass=user)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
[void]$searcher.PropertiesToLoad.AddRange(@(
    'givenName', 'sn', 'title', 'department',
    'telephoneNumber', 'mail', 'physicalDeliveryOfficeName'
))
$searcher.PageSize = 1000
$results = $searcher.FindAll()

Write-Host "[2/5] Got $($results.Count) records. Processing..."

# ============================================================
# FILTER: service accounts
# ============================================================
function Test-IsServiceAccount {
    param([string]$Sn, [string]$Gn, [string]$Email)
    $name = "$Sn $Gn".ToLower()
    if ($name -match ([char]0x0422 + [char]0x0421 + [char]0x0414)) { return $true }
    $patterns = @(
        'robot', 'temp', 'admin', 'operator', 'test', 'service',
        'desktop', 'horizon', 'view', '1c', 'terminal',
        'kassa', 'proizvodstvo', 'revizor', 'manager_avtd'
    )
    foreach ($p in $patterns) {
        if ($name -like "*$p*") { return $true }
    }
    if ([string]::IsNullOrEmpty($Sn) -or [string]::IsNullOrEmpty($Gn)) { return $true }
    if ($Email -match '^(robot|temp|admin|test|service|desktop|kassa|proizvodstvo|revizor)') { return $true }
    return $false
}

# ============================================================
# STEP 3: Collect ALL unique departments from AD
# ============================================================
Write-Host "[3/5] Scanning departments..."

$allDepts = [System.Collections.ArrayList]::new()
foreach ($r in $results) {
    $p = $r.Properties
    $dept = if ($p.department) { $p.department[0].ToString().Trim() } else { '' }
    if (-not [string]::IsNullOrEmpty($dept)) {
        $exists = $false
        foreach ($d in $allDepts) { if ($d -eq $dept) { $exists = $true; break } }
        if (-not $exists) { [void]$allDepts.Add($dept) }
    }
}
$sortedDepts = $allDepts | Sort-Object
Write-Host "  Found $($sortedDepts.Count) unique departments" -ForegroundColor Gray

# ============================================================
# STEP 4: Load or create dept-filter.json
# ============================================================
Write-Host "[4/5] Loading department filter..."

$deptConfig = $null
if (Test-Path $configFile) {
    try {
        $raw = Get-Content $configFile -Raw -Encoding UTF8
        $deptConfig = $raw | ConvertFrom-Json
        Write-Host "  Config loaded: $($deptConfig.Count) rules" -ForegroundColor Gray
    } catch {
        Write-Host "  Config error, creating new" -ForegroundColor Yellow
        $deptConfig = $null
    }
}

# Build config from AD data (MERGE with existing — preserve user edits)
$existingMap = @{}
if ($null -ne $deptConfig) {
    foreach ($c in $deptConfig) {
        if ($null -ne $c.name) {
            $existingMap[$c.name] = [bool]$c.enabled
        }
    }
}

$newConfig = [System.Collections.ArrayList]::new()
foreach ($dept in $sortedDepts) {
    if ($existingMap.ContainsKey($dept)) {
        # Department exists in config — keep user's enabled/disabled choice
        [void]$newConfig.Add(@{ name = $dept; enabled = $existingMap[$dept] })
    } else {
        # New department from AD — default to enabled
        [void]$newConfig.Add(@{ name = $dept; enabled = $true })
    }
}

# Save config (only if changed)
$changed = $false
if ($null -ne $deptConfig -and $newConfig.Count -eq $deptConfig.Count) {
    for ($i = 0; $i -lt $newConfig.Count; $i++) {
        if ($newConfig[$i].name -ne $deptConfig[$i].name -or $newConfig[$i].enabled -ne [bool]$deptConfig[$i].enabled) {
            $changed = $true; break
        }
    }
} else { $changed = $true }

if ($changed) {
    $newConfig | ConvertTo-Json -Depth 3 | Set-Content $configFile -Encoding UTF8
    Write-Host "  Config updated: $configFile ($($newConfig.Count) departments)" -ForegroundColor Gray
} else {
    Write-Host "  Config unchanged ($($newConfig.Count) departments)" -ForegroundColor Gray
}

# Build lookup: name -> enabled
$deptEnabled = @{}
foreach ($c in $newConfig) { $deptEnabled[$c.name] = [bool]$c.enabled }

# ============================================================
# STEP 5: Process records
# ============================================================
Write-Host "[5/5] Building employee list..."

$employees = [System.Collections.ArrayList]::new()
$excluded = 0
$skippedDept = 0

foreach ($r in $results) {
    $p = $r.Properties
    $gn = if ($p.givenname)   { $p.givenname[0].ToString().Trim() }   else { '' }
    $sn = if ($p.sn)          { $p.sn[0].ToString().Trim() }          else { '' }
    $email = if ($p.mail)     { $p.mail[0].ToString().Trim() }        else { '' }

    if (Test-IsServiceAccount -Sn $sn -Gn $gn -Email $email) { $excluded++; continue }

    $title = if ($p.title)                     { $p.title[0].ToString().Trim() }                     else { '' }
    $dept  = if ($p.department)                { $p.department[0].ToString().Trim() }                else { '' }
    $phone = if ($p.telephonenumber)           { $p.telephonenumber[0].ToString().Trim() }           else { '' }
    $room  = if ($p.physicaldeliveryofficename){ $p.physicaldeliveryofficename[0].ToString().Trim() } else { '' }

    if ($email -ne '' -and $email -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $email = '' }

    # Skip disabled departments
    if (-not [string]::IsNullOrEmpty($dept) -and $deptEnabled.ContainsKey($dept) -and -not $deptEnabled[$dept]) {
        $skippedDept++
        continue
    }

    # Generate slug from department name
    $deptSlug = 'no-dept'
    if (-not [string]::IsNullOrEmpty($dept)) {
        $hash = 0
        foreach ($ch in $dept.ToCharArray()) {
            $hash = ($hash * 31 + [int]$ch) % 10000
        }
        $deptSlug = 'dept-' + [Math]::Abs($hash)
    }

    [void]$employees.Add(@{
        last = $sn; first = $gn; dept = $deptSlug; deptName = $dept
        role = $title; phone = $phone; email = $email; room = $room
    })
}

$results.Dispose()
$searcher.Dispose()

Write-Host "  Employees: $($employees.Count), excluded: $excluded, skipped(dept): $skippedDept"

# Sort: by dept name, then by last name; empty dept last
$sorted = $employees | Sort-Object { if ([string]::IsNullOrEmpty($_.deptName)) { 'zzz' } else { $_.deptName } }, { $_.last }

# ============================================================
# GENERATE JS
# ============================================================
Write-Host "  Writing to $outputFile..."

# Collect unique departments for the header
$deptSummary = @{}
foreach ($emp in $sorted) {
    $key = $emp.dept
    if (-not $deptSummary.ContainsKey($key)) {
        $deptSummary[$key] = @{ slug = $emp.dept; name = $emp.deptName; count = 0 }
    }
    $deptSummary[$key].count++
}

$sb = [System.Text.StringBuilder]::new()
[void]$sb.AppendLine('// Auto-generated by Export-Phonebook.ps1')
[void]$sb.AppendLine('// Date: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
[void]$sb.AppendLine('// Source: Active Directory (' + $dcServer + ')')
[void]$sb.AppendLine('// Employees: ' + $sorted.Count)
[void]$sb.AppendLine('// Departments: ' + $deptSummary.Count)
[void]$sb.AppendLine()
[void]$sb.AppendLine('var employees = [')

$i = 0
foreach ($emp in $sorted) {
    $i++
    $comma = if ($i -lt $sorted.Count) { ',' } else { '' }
    $sn2 = $emp.last -replace "'", "\\'"
    $gn2 = $emp.first -replace "'", "\\'"
    $dn2 = $emp.deptName -replace "'", "\\'"
    $rl2 = $emp.role -replace "'", "\\'"
    $ph2 = $emp.phone -replace "'", "\\'"
    $em2 = $emp.email -replace "'", "\\'"
    $rm2 = $emp.room -replace "'", "\\'"
    [void]$sb.AppendLine("      { last: '$sn2', first: '$gn2', dept: '$($emp.dept)', deptName: '$dn2', role: '$rl2', phone: '$ph2', email: '$em2', room: '$rm2' }$comma")
}

[void]$sb.AppendLine('    ];')

$tempFile = $outputFile + '.tmp'
[System.IO.File]::WriteAllText($tempFile, $sb.ToString(), [System.Text.UTF8Encoding]::new($false))
if (Test-Path $outputFile) { Remove-Item $outputFile -Force }
Rename-Item $tempFile $outputFile

Write-Host ""
Write-Host "========================================"
Write-Host "  DONE!"
Write-Host "  File: $outputFile"
Write-Host "  Employees: $($sorted.Count)"
Write-Host "  Departments: $($deptSummary.Count)"
Write-Host "  Date: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')"
Write-Host "========================================"

} catch {
    Write-Host ""
    Write-Host "========================================"
    Write-Host "  ERROR!"
    Write-Host "========================================"
    Write-Host ""
    Write-Host $_.Exception.Message
}

Write-Host ""
Write-Host "Press Enter to close..."
Read-Host
