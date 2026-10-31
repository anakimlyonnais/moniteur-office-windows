Option Explicit
' MONITEUR OFFICE + WINDOWS V4 - un seul fichier a deposer sur le Bureau.
' VERSION_APP: 4.0.0
' Lancement manuel : 1 seule fenetre WPF, analyse immediate et journal visible.
' Surveillance : tache planifiee sans elevation et sans console CMD/PowerShell.
' Remediation : uniquement apres accord explicite, avec UAC si necessaire.
' Le script copie sa propre version dans LocalAppData pour fiabiliser la tache.
'
Dim fso, sh, root, installed, psFile, src, psExe, selectedMode, target
Dim q, cmd, result, input, output, line, copying, found
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
q = Chr(34)
src = WScript.ScriptFullName
root = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\ControleLicencesOfficeWindows"
installed = root & "\Moniteur-Licences.vbs"
psFile = root & "\Moniteur-Licences.ps1"
psExe = sh.ExpandEnvironmentStrings("%WINDIR%") & "\System32\WindowsPowerShell\v1.0\powershell.exe"
selectedMode = "Gui"
target = "Windows"

If WScript.Arguments.Count > 0 Then
    Select Case LCase(Trim(WScript.Arguments(0)))
        Case "/background": selectedMode = "Background"
        Case "/repair"
            selectedMode = "Repair"
            If WScript.Arguments.Count > 1 Then
                If LCase(WScript.Arguments(1)) = "office" Then target = "Office"
                If LCase(WScript.Arguments(1)) = "windows" Then target = "Windows"
            End If
        Case Else
            MsgBox "Argument non reconnu. Utilisez simplement le double-clic.", vbExclamation, "Moniteur licences V4"
            WScript.Quit 2
    End Select
End If

Function ReadVersion(p)
    On Error Resume Next
    Dim handle, i, ln
    ReadVersion = "0.0.0"
    If Not fso.FileExists(p) Then Exit Function
    Set handle = fso.OpenTextFile(p, 1, False)
    For i = 1 To 18
        If handle.AtEndOfStream Then Exit For
        ln = Trim(handle.ReadLine)
        If Left(ln, 15) = "' VERSION_APP: " Then
            ReadVersion = Trim(Mid(ln, 16))
            Exit For
        End If
    Next
    handle.Close
    Err.Clear
End Function

Function CompareVersions(a,b)
    On Error Resume Next
    Dim ar, br, i, va, vb
    ar = Split(a, ".")
    br = Split(b, ".")
    CompareVersions = 0
    For i = 0 To 2
        va = 0
        vb = 0
        If UBound(ar) >= i Then va = CLng(ar(i))
        If UBound(br) >= i Then vb = CLng(br(i))
        If va > vb Then CompareVersions = 1 : Exit Function
        If va < vb Then CompareVersions = -1 : Exit Function
    Next
    Err.Clear
End Function

' A desktop bootstrap must never downgrade a newer installed version.
If LCase(src) <> LCase(installed) And fso.FileExists(installed) Then
    If CompareVersions(ReadVersion(installed), ReadVersion(src)) >= 0 Then
        Dim forwardArgs, ix, callInstalled
        forwardArgs = ""
        For ix = 0 To WScript.Arguments.Count - 1
            forwardArgs = forwardArgs & " " & q & WScript.Arguments(ix) & q
        Next
        callInstalled = q & WScript.FullName & q & " " & q & installed & q & forwardArgs
        sh.Run callInstalled, 0, False
        WScript.Quit 0
    End If
End If

On Error Resume Next
If Not fso.FolderExists(root) Then fso.CreateFolder root
If LCase(src) <> LCase(installed) Then fso.CopyFile src, installed, True
If Err.Number <> 0 Then
    If selectedMode = "Gui" Then MsgBox "Impossible de preparer le moniteur : " & Err.Description, vbCritical, "Moniteur licences V4"
    WScript.Quit 3
End If
Err.Clear
Set input = fso.OpenTextFile(installed, 1, False)
Set output = fso.CreateTextFile(psFile, True, False)
If Err.Number <> 0 Then
    If selectedMode = "Gui" Then MsgBox "Impossible de charger le moteur d'analyse : " & Err.Description, vbCritical, "Moniteur licences V4"
    WScript.Quit 4
End If
copying = False
found = False
Do While Not input.AtEndOfStream
    line = input.ReadLine
    If line = "'@@BEGIN_POWERSHELL" Then
        copying = True
        found = True
    ElseIf line = "'@@END_POWERSHELL" Then
        Exit Do
    ElseIf copying And Left(line, 1) = "'" Then
        output.WriteLine Mid(line, 2)
    End If
Loop
input.Close
output.Close
If Err.Number <> 0 Or Not found Then
    If selectedMode = "Gui" Then MsgBox "Fichier incomplet : moteur PowerShell non extrait.", vbCritical, "Moniteur licences V4"
    WScript.Quit 5
End If
On Error GoTo 0

cmd = q & psExe & q & " -NoLogo -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass" & _
      " -File " & q & psFile & q & " -Mode " & selectedMode & " -Target " & target & _
      " -EntryFile " & q & installed & q
On Error Resume Next
Err.Clear
result = sh.Run(cmd, 0, (selectedMode <> "Gui"))
If Err.Number <> 0 And selectedMode = "Gui" Then
    MsgBox "Lancement impossible : " & Err.Description, vbCritical, "Moniteur licences V4"
End If
On Error GoTo 0
WScript.Quit 0

'@@BEGIN_POWERSHELL
'param(
'    [ValidateSet('Gui','Scan','Background','Repair','InstallUpdate')][string]$Mode = 'Gui',
'    [ValidateSet('Windows','Office')][string]$Target = 'Windows',
'    [string]$ReportFile = '',
'    [string]$ProgressFile = '',
'    [string]$EntryFile = '',
'    [switch]$UseReport
')
'
'$ErrorActionPreference = 'Stop'
'$Root = Join-Path $env:LOCALAPPDATA 'ControleLicencesOfficeWindows'
'$TaskName = 'Controle licences Office et Windows'
'$InstalledEntry = Join-Path $Root 'Moniteur-Licences.vbs'
'$LastReport = Join-Path $Root 'last-report-v3.json'
'$LastAlert = Join-Path $Root 'last-alert-v3.json'
'$RepairResult = Join-Path $Root 'repair-result-v3.json'
'$History = Join-Path $Root 'journal-v3.txt'
'$PSExe = Join-Path $PSHOME 'powershell.exe'
'$AppVersion = '4.0.0'
'$UpdateFeed = 'https://raw.githubusercontent.com/anakimlyonnais/moniteur-office-windows/main/updates/latest.json'
'$PublisherModulus = 'xS1lG4l6U3dONy9r1keZf3gKER7L/NzNf9SZue+gcPWl/9V9VxAf1uzdSXb40lTRZK7ZCuZa53M+7EkgZuJOsdASyaCrAqrReWgBGbaZF7WrdF+Q4ynbHO/ne+3O43PSrNumt0OXd/3ZGqw7jKUwtMdmT+Dxd2zaAFsqcGFUo9eGAAAeB6Kwk/BWSVgpr2JT2rmXdpzvtobTK2GgPFfh787enxvKmTKyV8v0Nk0dxYNzi3pp/dJ2WcdnBGLi/yeNsmTczrBiO+J/jTY6zE1dWQ++fU6krdju1JJoN90sZJLySG+mF0PYJuMrvxBcTBZYMTkL2EGp1ElNPm3p2KL9+KQbIdegQUEycJLX3QiLgb+AazflwatKYQPgJUlGzFhwyjGriWn8HtEXghGASwASGIF/k9m2Y0eomRKCAqxUQJgvHlepY1Y4RY84SuT2NQge33+I6KVqwNYO19tVsaeleFHvKT9qNjI2Ig9Pz+x9wVUTJo8xsxyctqHOu797kvjZ'
'$PublisherExponent = 'AQAB'
'$UpdateSnooze = Join-Path $Root 'update-snooze-v4.json'
'$UpdateResult = Join-Path $Root 'update-result-v4.json'
'if (-not (Test-Path -LiteralPath $Root)) { [void](New-Item -ItemType Directory -Path $Root -Force) }
'
'function Add-History([string]$Text) {
'    try { Add-Content -LiteralPath $History -Encoding UTF8 -Value ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Text) } catch { }
'}
'function Step([string]$Text) {
'    if ($ProgressFile) {
'        try { Add-Content -LiteralPath $ProgressFile -Encoding UTF8 -Value ('[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Text) } catch { }
'    }
'}
'function Read-Json([string]$Path) {
'    if (-not (Test-Path -LiteralPath $Path)) { return $null }
'    try { return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
'}
'function Save-Json([string]$Path, $Data) {
'    $tmp = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
'    try {
'        $Data | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $tmp -Encoding UTF8
'        Move-Item -LiteralPath $tmp -Destination $Path -Force
'    } finally {
'        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
'    }
'}
'function Get-Status([int]$Code) {
'    switch ($Code) { 1 {'ACTIVE'} 0 {'INACTIVE'} 5 {'INACTIVE'} 2 {'ATTENTION'} 3 {'ATTENTION'} 4 {'ATTENTION'} 6 {'ATTENTION'} default {'INCONNU'} }
'}
'function Get-Expiry($Product) {
'    try {
'        $minutes = [double]$Product.GracePeriodRemaining
'        if ($minutes -gt 0) { return ('{0} (estimation)' -f (Get-Date).AddMinutes($minutes).ToString('dd/MM/yyyy HH:mm')) }
'        if ([int]$Product.LicenseStatus -eq 1) { return 'Pas d''echeance indiquee (licence permanente possible)' }
'    } catch { }
'    return 'Non disponible'
'}
'function Get-WindowsInfo {
'    $res = [ordered]@{ Edition = 'Non disponible'; Build = 'Non disponible'; State='INCONNU'; Channel='Non disponible'; Key='Non disponible'; Expiry='Non disponible'; Last='Non disponible'; Reason=''; Detail='' }
'    try {
'        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
'        $res.Edition = [string]$os.Caption
'        $res.Build = ('{0} / {1} / build {2}' -f $os.Version, $os.OSArchitecture, $os.BuildNumber)
'        $products = @(Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" -ErrorAction Stop |
'            Where-Object { -not $_.LicenseIsAddon -and $_.Name -match 'Windows' })
'        if ($products.Count -eq 0) {
'            $res.Detail = 'Aucune licence Windows identifiable par WMI.'
'        } else {
'            $p = $products | Sort-Object @{Expression={if ([int]$_.LicenseStatus -eq 1) {0} else {1}}}, @{Expression={if ($_.Description -match 'Retail|OEM|VOLUME') {0} else {1}}} | Select-Object -First 1
'            $res.State = Get-Status ([int]$p.LicenseStatus)
'            if ($p.PartialProductKey) { $res.Key = ('*****-{0}' -f [string]$p.PartialProductKey) }
'            if ($p.Description) { $res.Channel = [string]$p.Description }
'            $res.Expiry = Get-Expiry $p
'            if ($p.LicenseStatusReason -and $res.State -ne 'ACTIVE') { $res.Reason = ('0x{0:X8}' -f [uint32]$p.LicenseStatusReason) }
'            $res.Detail = 'Statut etabli avec SoftwareLicensingProduct (Windows).'
'        }
'    } catch {
'        $res.Detail = 'WMI indisponible : ' + $_.Exception.Message
'    }
'    # Windows does not expose a reliable last-activation timestamp via this API.
'    return [pscustomobject]$res
'}
'function Office-Roots {
'    $list = New-Object 'System.Collections.Generic.List[string]'
'    foreach ($rootPath in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
'        if (-not $rootPath) { continue }
'        foreach ($relative in @('Microsoft Office\root\Office16','Microsoft Office\Office16','Microsoft Office\Office15')) {
'            $p = Join-Path $rootPath $relative
'            if ((Test-Path -LiteralPath $p) -and (-not $list.Contains($p))) { $list.Add($p) }
'        }
'    }
'    return @($list.ToArray())
'}
'function Call-Cscript([string]$ScriptPath, [string]$Parameter) {
'    $start = New-Object System.Diagnostics.ProcessStartInfo
'    $start.FileName = Join-Path $env:WINDIR 'System32\cscript.exe'
'    $start.Arguments = ('//Nologo "{0}" {1}' -f $ScriptPath, $Parameter)
'    $start.UseShellExecute = $false
'    $start.CreateNoWindow = $true
'    $start.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
'    $start.RedirectStandardOutput = $true
'    $start.RedirectStandardError = $true
'    $proc = New-Object System.Diagnostics.Process
'    $proc.StartInfo = $start
'    try {
'        [void]$proc.Start()
'        $stdout = $proc.StandardOutput.ReadToEnd()
'        $stderr = $proc.StandardError.ReadToEnd()
'        if (-not $proc.WaitForExit(60000)) { $proc.Kill(); throw 'Delai diagnostic depasse (60 s).' }
'        return [pscustomobject]@{ Text=($stdout + "`n" + $stderr); ExitCode=$proc.ExitCode }
'    } finally { $proc.Dispose() }
'}
'function Get-OfficeInfo {
'    $res = [ordered]@{ Edition='Non detectee'; Version='Non disponible'; State='ABSENT'; Type='Non disponible'; Key='Non disponible'; Expiry='Non disponible'; Last='Non disponible'; Apps='Aucune suite Office'; Detail=''; Repairable=$false; OSPP='' }
'    $c2r = $null
'    try { $c2r = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop } catch { }
'    if ($c2r -and $c2r.ProductReleaseIds) { $res.Edition = [string]$c2r.ProductReleaseIds }
'    if ($c2r -and $c2r.VersionToReport) { $res.Version = [string]$c2r.VersionToReport }
'    if ($c2r -and $c2r.Platform) { $res.Version += (' / {0}' -f [string]$c2r.Platform) }
'    $roots = @(Office-Roots)
'    $found = @{}
'    foreach ($path in $roots) {
'        foreach ($name in @('WINWORD','EXCEL','POWERPNT','OUTLOOK','ONENOTE','MSACCESS','MSPUB')) {
'            $exe = Join-Path $path ($name + '.EXE')
'            if (-not $found.ContainsKey($name) -and (Test-Path -LiteralPath $exe)) {
'                $found[$name] = $exe
'                if ($res.Version -eq 'Non disponible') {
'                    try { $res.Version = [string](Get-Item -LiteralPath $exe).VersionInfo.ProductVersion } catch { }
'                }
'            }
'        }
'    }
'    if ($found.Count -eq 0) { $res.Detail = 'Word, Excel et PowerPoint non reperes dans les chemins Office habituels.'; return [pscustomobject]$res }
'    $res.Apps = (($found.Keys | Sort-Object) -join ', ')
'    if (-not ($found.ContainsKey('WINWORD') -and $found.ContainsKey('EXCEL'))) {
'        $res.State = 'PARTIEL'
'        $res.Detail = 'Installation Office partielle (Word et Excel non trouves ensemble).'
'        return [pscustomobject]$res
'    }
'    $res.State = 'INCONNU'
'    if ($res.Edition -eq 'Non detectee') { $res.Edition = 'Microsoft Office (edition non identifiee)' }
'    $is365 = ($res.Edition -match '(?i)O365|M365|365|Subscription')
'    $res.Type = if ($is365) { 'Microsoft 365 / compte' } else { 'Licence produit ou volume (a verifier)' }
'    Step 'Office installe. Recherche du canal de licence...'
'    $vnextLicensed=$false; $vnextInactive=$false; $volumeLicensed=$false; $volumeInactive=$false
'    foreach ($path in $roots) {
'        $vnext = Join-Path $path 'vnextdiag.ps1'
'        if (Test-Path -LiteralPath $vnext) {
'            try {
'                $out = (& $vnext -action list 2>&1 | Out-String)
'                if ($out -match '(?im)^\s*(?:State|License State|License Status|Activation State)\s*:\s*(?:Licensed|Active)\b') { $vnextLicensed=$true }
'                if ($out -match '(?im)^\s*(?:State|License State|License Status|Activation State)\s*:\s*(?:Unlicensed|Expired|Notification|Disabled|Not Licensed)\b') { $vnextInactive=$true }
'            } catch { $res.Detail += ' vnextdiag inconclusif.' }
'        }
'        $ospp = Join-Path $path 'OSPP.VBS'
'        if (Test-Path -LiteralPath $ospp) {
'            try {
'                $result = Call-Cscript $ospp '/dstatus'
'                $blocks = [regex]::Split([string]$result.Text, '(?m)^\s*-{10,}\s*$')
'                foreach ($block in $blocks) {
'                    if ($block -notmatch '(?im)^\s*LICENSE NAME\s*:\s*(?<name>[^\r\n]+)') { continue }
'                    $licenseName = [string]$Matches.name
'                    if ($licenseName -match '(?i)Project|Visio') { continue }
'                    if ($licenseName -notmatch '(?i)Office') { continue }
'                    if ($block -match '(?im)LICENSE DESCRIPTION\s*:\s*([^\r\n]+)') { $res.Type = [string]$Matches[1] }
'                    if ($block -match '(?im)Last 5 characters of installed product key\s*:\s*([A-Z0-9]{5})') { $res.Key = '*****-' + [string]$Matches[1] }
'                    if ($block -match '(?im)REMAINING GRACE\s*:\s*(\d+)\s*days?') {
'                        $days = [int]$Matches[1]
'                        if ($days -gt 0) { $res.Expiry = ('{0} (approx.)' -f (Get-Date).AddDays($days).ToString('dd/MM/yyyy')) }
'                    }
'                    if ($block -match '(?im)LICENSE STATUS\s*:\s*---LICENSED---') { $volumeLicensed=$true; $res.OSPP=$ospp; $res.Repairable=$true }
'                    elseif ($block -match '(?im)LICENSE STATUS\s*:\s*---([A-Z_]+)---') { $volumeInactive=$true; $res.OSPP=$ospp; $res.Repairable=$true }
'                }
'            } catch { $res.Detail += ' ospp inconclusif.' }
'        }
'    }
'    if ($is365) {
'        if ($vnextLicensed) { $res.State='ACTIVE'; $res.Type='Microsoft 365 / abonnement' }
'        elseif ($vnextInactive) { $res.State='INACTIVE'; $res.Type='Microsoft 365 / abonnement' }
'        elseif ($volumeLicensed) { $res.State='ACTIVE' }
'        elseif ($volumeInactive) { $res.State='INACTIVE' }
'    } else {
'        if ($volumeLicensed) { $res.State='ACTIVE' }
'        elseif ($volumeInactive) { $res.State='INACTIVE' }
'        elseif ($vnextLicensed) { $res.State='ACTIVE'; $res.Type='Microsoft 365 / compte' }
'        elseif ($vnextInactive) { $res.State='INACTIVE'; $res.Type='Microsoft 365 / compte' }
'    }
'    if ($is365 -and $res.Expiry -eq 'Non disponible') { $res.Expiry = 'Non exposee par vnextdiag' }
'    if ($res.State -eq 'INCONNU') { $res.Detail += ' Licence non confirmee; ne signifie pas inactive.' }
'    else { $res.Detail += ' Statut identifie avec les outils Microsoft Office.' }
'    if (-not $res.Detail) { $res.Detail = 'Controle Office termine.' }
'    return [pscustomobject]$res
'}
'function Get-OfficeProcesses {
'    return @(Get-Process -Name WINWORD,EXCEL,POWERPNT,OUTLOOK,ONENOTE,MSACCESS,MSPUB,MSPROJECT,VISIO -ErrorAction SilentlyContinue)
'}
'function Get-Advice($W, $O) {
'    if ($W.State -eq 'INACTIVE') { return [pscustomobject]@{ Action='WindowsRepair'; Message='Windows semble desactive. Tenter la reactivation Microsoft avec la licence deja installee. Une confirmation et l''UAC seront demandes.'; Button='TENTER REACTIVATION WINDOWS'; Alert=$true } }
'    if ($O.State -eq 'ABSENT' -or $O.State -eq 'PARTIEL') { return [pscustomobject]@{ Action='OfficeInstall'; Message='Suite Office absente ou incomplete. Ouvrir la page Microsoft pour installer l''edition associee a votre licence.'; Button='INSTALLER OFFICE (MICROSOFT)'; Alert=$true } }
'    if ($O.State -eq 'INACTIVE' -and $O.Repairable -and $O.Type -match '(?i)VOLUME') { return [pscustomobject]@{ Action='OfficeRepair'; Message='Office a une licence en volume inactive. Fermer les applications Office, puis demander une reactivation officielle avec la cle deja installee.'; Button='TENTER REACTIVATION OFFICE'; Alert=$true } }
'    if ($O.State -eq 'INACTIVE') { return [pscustomobject]@{ Action='OfficeAccount'; Message='Office est inactif. Pour Microsoft 365 ou une licence personnelle, reconnectez le compte autorise dans Word > Fichier > Compte.'; Button='OUVRIR WORD / COMPTE'; Alert=$true } }
'    if ($W.State -eq 'ATTENTION') { return [pscustomobject]@{ Action='WindowsRepair'; Message='Windows est en periode de grace. Verification et reactivation officielle conseillees avant la date indiquee.'; Button='TENTER REACTIVATION WINDOWS'; Alert=$true } }
'    if ($O.State -eq 'INCONNU') { return [pscustomobject]@{ Action='OfficeAccount'; Message='Office est installe mais sa licence ne peut pas etre confirmee avec certitude. Verifiez dans Word > Fichier > Compte.'; Button='OUVRIR WORD / COMPTE'; Alert=$false } }
'    if ($W.State -eq 'INCONNU') { return [pscustomobject]@{ Action='WindowsSettings'; Message='Etat de Windows non confirme. Examiner la page officielle des parametres d''activation.'; Button='PARAMETRES ACTIVATION WINDOWS'; Alert=$false } }
'    return [pscustomobject]@{ Action='None'; Message='Windows et Office sont detectes comme actives. Aucun changement de licence necessaire.'; Button='AUCUNE ACTION NECESSAIRE'; Alert=$false }
'}
'function Get-UpdateInfo {
'    $info = [ordered]@{ Available=$false; Verified=$false; Version=$AppVersion; Notes=''; Url=''; Sha256=''; Status='Verification indisponible' }
'    try {
'        # TLS is only used to retrieve data; the offline RSA signature is the trust anchor.
'        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
'        $manifest = Invoke-RestMethod -Uri $UpdateFeed -Method Get -UseBasicParsing -TimeoutSec 12 -ErrorAction Stop
'        $version = [string]$manifest.version
'        $sha = [string]$manifest.sha256
'        $url = [string]$manifest.url
'        $notes = [string]$manifest.notes
'        $signature = [string]$manifest.signature
'        if ($version -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}$') { throw 'Version du manifeste invalide.' }
'        if ($sha -notmatch '^[a-fA-F0-9]{64}$') { throw 'Empreinte SHA256 invalide.' }
'        if ($notes.Length -gt 1000) { throw 'Notes de version trop longues.' }
'        if ($url -notmatch '^https://raw\.githubusercontent\.com/anakimlyonnais/moniteur-office-windows/[A-Za-z0-9_.\-/]+\.vbs$') { throw 'URL hors depot autorise.' }
'        $payload = "$version`n$url`n$($sha.ToLowerInvariant())`n$notes`n"
'        [byte[]]$data = [System.Text.Encoding]::UTF8.GetBytes($payload)
'        [byte[]]$sig = [Convert]::FromBase64String($signature)
'        $param = New-Object System.Security.Cryptography.RSAParameters
'        $param.Modulus = [Convert]::FromBase64String($PublisherModulus)
'        $param.Exponent = [Convert]::FromBase64String($PublisherExponent)
'        $rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
'        try {
'            $rsa.ImportParameters($param)
'            if (-not $rsa.VerifyData($data, 'SHA256', $sig)) { throw 'Signature du manifeste invalide.' }
'        } finally { $rsa.Dispose() }
'        $info.Verified=$true
'        $info.Version=$version
'        $info.Notes=$notes
'        $info.Url=$url
'        $info.Sha256=$sha.ToLowerInvariant()
'        if ([version]$version -gt [version]$AppVersion) {
'            $info.Available=$true
'            $info.Status='Nouvelle version signee disponible.'
'        } else {
'            $info.Status='Version actuelle a jour.'
'        }
'    } catch {
'        $info.Status='Verification impossible (aucune installation) : ' + $_.Exception.Message
'    }
'    return [pscustomobject]$info
'}
'function Install-VerifiedUpdate {
'    $result = [ordered]@{ Success=$false; Message='Mise a jour non effectuee'; Version=$AppVersion; Timestamp=(Get-Date).ToString('o') }
'    $temp = Join-Path $Root ('update-' + [guid]::NewGuid().ToString('N') + '.vbs')
'    $backup = Join-Path $Root 'Moniteur-Licences-backup-v4.vbs'
'    $backedUp = $false
'    $replaced = $false
'    try {
'        $up = Get-UpdateInfo
'        if (-not ($up.Available -and $up.Verified)) { throw ('Aucune nouvelle version signee : ' + $up.Status) }
'        # Downloaded programs cannot run or replace an installation until hash and signature are checked.
'        Invoke-WebRequest -Uri $up.Url -UseBasicParsing -OutFile $temp -TimeoutSec 40 -ErrorAction Stop
'        $size = (Get-Item -LiteralPath $temp).Length
'        if ($size -lt 20000 -or $size -gt 250000) { throw 'Taille du fichier de mise a jour non conforme.' }
'        $hash = (Get-FileHash -LiteralPath $temp -Algorithm SHA256).Hash.ToLowerInvariant()
'        if ($hash -ne $up.Sha256) { throw 'Le fichier ne correspond pas a la signature de version.' }
'        if (-not (Test-Path -LiteralPath $InstalledEntry)) { throw 'Installation locale introuvable.' }
'        Copy-Item -LiteralPath $InstalledEntry -Destination $backup -Force
'        $backedUp = $true
'        $replaced = $true
'        Copy-Item -LiteralPath $temp -Destination $InstalledEntry -Force
'        $hashAfter = (Get-FileHash -LiteralPath $InstalledEntry -Algorithm SHA256).Hash.ToLowerInvariant()
'        if ($hashAfter -ne $up.Sha256) { throw 'Controle apres copie echoue.' }
'        $result.Success = $true
'        $result.Version = $up.Version
'        $result.Message = ('Version {0} verifiee et installee. Demarrage dans la meme interface.' -f $up.Version)
'        Add-History ('Mise a jour approuvee et verifiee : ' + $up.Version)
'    } catch {
'        if ($replaced -and $backedUp -and (Test-Path -LiteralPath $backup)) {
'            try { Copy-Item -LiteralPath $backup -Destination $InstalledEntry -Force } catch { }
'        }
'        $result.Message = 'Mise a jour annulee : ' + $_.Exception.Message
'        Add-History $result.Message
'    } finally {
'        if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
'        Save-Json $UpdateResult ([pscustomobject]$result)
'    }
'}
'function Run-Scan {
'    Step 'Initialisation du controle local...'
'    Step 'Windows : version, architecture, edition, licence...'
'    $w = Get-WindowsInfo
'    Step ('Windows : resultat {0}.' -f $w.State)
'    Step 'Office : recherche Word, Excel, PowerPoint, version...'
'    $o = Get-OfficeInfo
'    Step ('Office : resultat {0}.' -f $o.State)
'    Step 'Applications ouvertes : verification sans fermeture...'
'    $open = @((Get-OfficeProcesses) | Select-Object -ExpandProperty ProcessName -Unique)
'    if ($open.Count) { Step ('Office en cours : ' + ($open -join ', ')) } else { Step 'Aucune application Office ouverte.' }
'    $advice = Get-Advice $w $o
'    Step ('Etape conseillee : ' + $advice.Action)
'    Step 'Mises a jour : verification du manifeste signe...'
'    $update = Get-UpdateInfo
'    Step ('Mises a jour : ' + $update.Status)
'    $result = [pscustomobject]@{ Timestamp=(Get-Date).ToString('o'); Windows=$w; Office=$o; OpenOffice=$open; Advice=$advice; Update=$update }
'    if ($ReportFile) { Save-Json $ReportFile $result }
'    Save-Json $LastReport $result
'    Add-History ('Controle : Windows={0} Office={1} Action={2}' -f $w.State, $o.State, $advice.Action)
'    Step 'Diagnostic termine. Aucune modification de licence effectuee.'
'    return $result
'}
'function Get-Monitor {
'    $result = [ordered]@{ Enabled=$false; Frequency='weekly'; Label='DESACTIVEE'; Legacy=$false }
'    try {
'        $svc = New-Object -ComObject 'Schedule.Service'
'        $svc.Connect()
'        $task = $svc.GetFolder('\').GetTask($TaskName)
'        $result.Enabled = [bool]$task.Enabled
'        if ($task.Definition.Triggers.Count -gt 0) {
'            $type = [int]$task.Definition.Triggers.Item(1).Type
'            if ($type -eq 2) { $result.Frequency = 'daily' } elseif ($type -eq 3) { $result.Frequency = 'weekly' }
'        }
'        if ($task.Definition.Actions.Count -gt 0) {
'            $a = [string]$task.Definition.Actions.Item(1).Arguments
'            $result.Legacy = ($a -notmatch '(?i)/background')
'        }
'        if ($result.Enabled) { $result.Label = if ($result.Frequency -eq 'daily') { 'QUOTIDIENNE' } else { 'HEBDOMADAIRE' } }
'        if ($result.Legacy -and $result.Enabled) { $result.Label += ' - ancienne version a migrer' }
'    } catch {
'        $msg = [string]$_.Exception.Message
'        $hr = $_.Exception.HResult
'        if ($hr -ne -2147024894 -and $msg -notmatch '0x80070002|cannot find|introuvable|specified file') {
'            $result.Label = 'ETAT INDISPONIBLE - ' + $msg
'        }
'    }
'    return [pscustomobject]$result
'}
'function Set-Monitor([bool]$Enabled, [string]$Frequency) {
'    $svc = New-Object -ComObject 'Schedule.Service'
'    $svc.Connect()
'    $folder = $svc.GetFolder('\')
'    if (-not $Enabled) {
'        try { $folder.DeleteTask($TaskName,0) } catch { if ($_.Exception.Message -notmatch '0x80070002|cannot find|introuvable') { throw } }
'        Add-History 'Surveillance desactivee a la demande de l''utilisateur.'
'        return
'    }
'    if (-not (Test-Path -LiteralPath $InstalledEntry)) { throw 'Lanceur installe introuvable. Relancez le fichier sur votre Bureau.' }
'    $d = $svc.NewTask(0)
'    $d.RegistrationInfo.Description = 'Moniteur Office et Windows V4 : verification non invasive. Alertes uniquement si anomalie certaine.'
'    $d.Settings.Enabled = $true
'    $d.Settings.Hidden = $false
'    $d.Settings.StartWhenAvailable = $true
'    $d.Settings.DisallowStartIfOnBatteries = $false
'    $d.Settings.StopIfGoingOnBatteries = $false
'    $d.Settings.ExecutionTimeLimit = 'PT8M'
'    $d.Principal.LogonType = 3
'    $d.Principal.RunLevel = 0
'    $d.Principal.UserId = ('{0}\{1}' -f $env:USERDOMAIN,$env:USERNAME)
'    if ($Frequency -eq 'daily') {
'        $t = $d.Triggers.Create(2)
'        $t.DaysInterval = 1
'    } else {
'        $t = $d.Triggers.Create(3)
'        $t.DaysOfWeek = 1
'        $t.WeeksInterval = 1
'    }
'    $t.StartBoundary = (Get-Date).Date.AddHours(12).ToString('yyyy-MM-ddTHH:mm:ss')
'    $t.Enabled = $true
'    $a = $d.Actions.Create(0)
'    $a.Path = Join-Path $env:WINDIR 'System32\wscript.exe'
'    $a.Arguments = ('"{0}" /background' -f $InstalledEntry)
'    $a.WorkingDirectory = $Root
'    [void]$folder.RegisterTaskDefinition($TaskName,$d,6,$null,$null,3)
'    Add-History ('Surveillance activee : ' + $Frequency)
'}
'function Run-Background {
'    $data = Run-Scan
'    $updateAlert = [bool]$data.Update.Available
'    if ($updateAlert) {
'        $snooze = Read-Json $UpdateSnooze
'        if ($snooze -and $snooze.Version -eq $data.Update.Version) {
'            try { if (((Get-Date) - [datetime]$snooze.Timestamp).TotalDays -lt 7) { $updateAlert=$false } } catch { }
'        }
'    }
'    if (-not [bool]$data.Advice.Alert -and -not $updateAlert) { return }
'    $fingerprint = ('Windows={0};Office={1};Action={2};Update={3}' -f $data.Windows.State,$data.Office.State,$data.Advice.Action,$(if ($updateAlert) {$data.Update.Version} else {'none'}))
'    $last = Read-Json $LastAlert
'    if ($last -and $last.Fingerprint -eq $fingerprint) {
'        try { if (((Get-Date) - [datetime]$last.Timestamp).TotalDays -lt 7) { Add-History 'Alerte identique deja presente : aucune fenetre affichee.'; return } } catch { }
'    }
'    Save-Json $LastAlert ([pscustomobject]@{ Fingerprint=$fingerprint;Timestamp=(Get-Date).ToString('o') })
'    Add-History 'Anomalie confirmee : ouverture du tableau de bord pour information.'
'    # Only the branded GUI opens on an actionable anomaly; no CMD or console window.
'    $start = New-Object System.Diagnostics.ProcessStartInfo
'    $start.FileName = $PSExe
'    $start.Arguments = ('-NoLogo -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Mode Gui -UseReport -EntryFile "{1}"' -f $PSCommandPath,$InstalledEntry)
'    $start.UseShellExecute = $false
'    $start.CreateNoWindow = $true
'    $start.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
'    [void][System.Diagnostics.Process]::Start($start)
'}
'function Run-Repair {
'    $out = [ordered]@{ Target=$Target; Success=$false; Message=''; Timestamp=(Get-Date).ToString('o') }
'    try {
'        $id=[Security.Principal.WindowsIdentity]::GetCurrent()
'        $isAdmin=([Security.Principal.WindowsPrincipal]::new($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
'        if (-not $isAdmin) { throw 'Droits administrateur requis.' }
'        $active = @(Get-OfficeProcesses)
'        if ($active.Count -gt 0) { throw 'Office a ete rouvert : intervention annulee. Sauvegardez et fermez les applications Office.' }
'        if ($Target -eq 'Windows') {
'            if ((Get-WindowsInfo).State -eq 'ACTIVE') {
'                $out.Success=$true
'                $out.Message='Windows est deja active. Aucune intervention effectuee.'
'                Save-Json $RepairResult ([pscustomobject]$out)
'                return
'            }
'            $file = Join-Path $env:WINDIR 'System32\slmgr.vbs'
'            $result = Call-Cscript $file '/ato'
'        } else {
'            $office = Get-OfficeInfo
'            if ($office.State -ne 'INACTIVE' -or -not $office.Repairable -or -not $office.OSPP -or $office.Type -notmatch '(?i)VOLUME') {
'                throw 'Licence Office en volume inactive non confirmee : aucune modification.'
'            }
'            $result = Call-Cscript $office.OSPP '/act'
'        }
'        $out.Success = ($result.ExitCode -eq 0)
'        $out.Message = if ($out.Success) { 'Tentative officielle terminee. Un nouveau diagnostic va confirmer le statut.' } else { ('Echec de la tentative officielle, code ' + $result.ExitCode) }
'        Add-History ('Tentative d''activation officielle ' + $Target + ', code ' + $result.ExitCode)
'    } catch { $out.Message = $_.Exception.Message; Add-History ('Reparation interrompue ' + $Target + ' : ' + $out.Message) }
'    Save-Json $RepairResult ([pscustomobject]$out)
'}
'
'if ($Mode -eq 'Scan') {
'    try { [void](Run-Scan); exit 0 }
'    catch { Step ('ERREUR DIAGNOSTIC : ' + $_.Exception.Message); Add-History ('Erreur scan : ' + $_.Exception.Message); exit 2 }
'}
'if ($Mode -eq 'Background') {
'    try { Run-Background } catch { Add-History ('Erreur surveillance : ' + $_.Exception.Message) }
'    exit
'}
'if ($Mode -eq 'Repair') { Run-Repair; exit }
'if ($Mode -eq 'InstallUpdate') { Install-VerifiedUpdate; exit }
'
'# One WPF window; all normal diagnostics, logs, recommendations and preferences stay here.
'try {
'Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
'[xml]$xml = @'
'<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
' Title="CONTROLE OFFICE + WINDOWS  |  MONITEUR V4" Height="850" Width="1080" MinHeight="630" MinWidth="840"
' WindowStartupLocation="CenterScreen" Background="#0A0D12" Foreground="#EBF3F3" FontFamily="Segoe UI">
' <Window.Resources>
'  <Style TargetType="Button">
'   <Setter Property="Background" Value="#21342A"/><Setter Property="Foreground" Value="#8DF3AD"/>
'   <Setter Property="BorderBrush" Value="#3C7758"/><Setter Property="BorderThickness" Value="1"/>
'   <Setter Property="FontFamily" Value="Consolas"/><Setter Property="FontWeight" Value="SemiBold"/>
'   <Setter Property="FontSize" Value="13"/><Setter Property="Padding" Value="14,9"/>
'   <Setter Property="Margin" Value="0,0,9,0"/><Setter Property="Cursor" Value="Hand"/>
'  </Style>
'  <Style TargetType="RadioButton"><Setter Property="Foreground" Value="#DAE8E5"/><Setter Property="Margin" Value="0,0,19,0"/><Setter Property="FontSize" Value="13"/></Style>
'  <Style TargetType="CheckBox"><Setter Property="Foreground" Value="#DAE8E5"/><Setter Property="FontSize" Value="14"/></Style>
' </Window.Resources>
' <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
'  <StackPanel Margin="25,21,25,19">
'   <Grid Margin="0,0,0,17">
'    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
'    <StackPanel>
'     <TextBlock FontFamily="Consolas" Foreground="#8DF3AD" FontSize="24" FontWeight="Bold" Text="&gt; CONTR&#212;LE DES LICENCES_"/>
'     <TextBlock Margin="1,6,0,0" Foreground="#8FA09C" FontSize="12" Text="MICROSOFT WINDOWS  /  OFFICE  |  DIAGNOSTIC LOCAL ET SURVEILLANCE"/>
'    </StackPanel>
'    <StackPanel Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Center">
'     <TextBlock x:Name="TopStatus" Foreground="#F5C879" FontFamily="Consolas" FontSize="13" TextAlignment="Right" Text="[ INITIALISATION ]"/>
'     <TextBlock x:Name="TopTime" Foreground="#7C918B" FontSize="12" TextAlignment="Right" Margin="0,7,0,0" Text="Aucune analyse terminee"/>
'    </StackPanel>
'   </Grid>
'   <Grid Margin="0,0,0,14">
'    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
'    <Border Grid.Column="0" Background="#131B22" BorderBrush="#293F34" BorderThickness="1" CornerRadius="7" Padding="16,14">
'     <StackPanel>
'      <DockPanel Margin="0,0,0,14">
'       <TextBlock Text="01 / WINDOWS" FontFamily="Consolas" FontSize="16" FontWeight="Bold" Foreground="#E0EDE6" DockPanel.Dock="Left"/>
'       <TextBlock x:Name="WinBadge" Text="ANALYSE..." Foreground="#F5C879" FontFamily="Consolas" FontWeight="Bold" FontSize="14" HorizontalAlignment="Right"/>
'      </DockPanel>
'      <StackPanel x:Name="WinFields"/>
'     </StackPanel>
'    </Border>
'    <Border Grid.Column="2" Background="#131B22" BorderBrush="#293F34" BorderThickness="1" CornerRadius="7" Padding="16,14">
'     <StackPanel>
'      <DockPanel Margin="0,0,0,14">
'       <TextBlock Text="02 / OFFICE" FontFamily="Consolas" FontSize="16" FontWeight="Bold" Foreground="#E0EDE6" DockPanel.Dock="Left"/>
'       <TextBlock x:Name="OfficeBadge" Text="ANALYSE..." Foreground="#F5C879" FontFamily="Consolas" FontWeight="Bold" FontSize="14" HorizontalAlignment="Right"/>
'      </DockPanel>
'      <StackPanel x:Name="OfficeFields"/>
'     </StackPanel>
'    </Border>
'   </Grid>
'   <Border Background="#151F1B" BorderBrush="#315743" BorderThickness="1" CornerRadius="7" Padding="16,12" Margin="0,0,0,12">
'    <StackPanel>
'     <TextBlock Foreground="#8DF3AD" FontFamily="Consolas" FontWeight="Bold" FontSize="13" Text="&gt;&gt; PROCHAINE &#201;TAPE LOGIQUE"/>
'     <TextBlock x:Name="AdviceText" Text="Analyse initiale en cours. Aucune modification effectu&#233;e." TextWrapping="Wrap" FontSize="13" Foreground="#E2ECE5" Margin="0,8,0,10"/>
'     <StackPanel Orientation="Horizontal">
'      <Button x:Name="ActionButton" Content="ANALYSE EN COURS..." IsEnabled="False"/>
'      <Button x:Name="RefreshButton" Content="REFAIRE LE CONTR&#212;LE" Background="#1D2830" Foreground="#D1DEE0" BorderBrush="#48606C"/>
'     </StackPanel>
'     <Border x:Name="ConfirmPanel" Background="#2A2922" BorderBrush="#987D42" BorderThickness="1" Padding="12" Margin="0,10,0,0" CornerRadius="4" Visibility="Collapsed">
'      <StackPanel>
'       <TextBlock x:Name="ConfirmText" TextWrapping="Wrap" Foreground="#FFE4B0" Margin="0,0,0,9"/>
'       <StackPanel Orientation="Horizontal">
'        <Button x:Name="YesButton" Content="[1] OUI" Background="#2B4434"/>
'        <Button x:Name="NoButton" Content="[0] NON" Background="#392626" Foreground="#F0CBCB" BorderBrush="#775353"/>
'       </StackPanel>
'      </StackPanel>
'     </Border>
'    </StackPanel>
'   </Border>
'   <Border Background="#131B22" BorderBrush="#315743" BorderThickness="1" CornerRadius="7" Padding="15,11" Margin="0,0,0,12">
'    <StackPanel>
'     <TextBlock Text="&gt; MISES A JOUR SIGNEES / CONFIRMATION REQUISE" Foreground="#CDEFE0" FontFamily="Consolas" FontWeight="Bold" FontSize="13" Margin="0,0,0,8"/>
'     <TextBlock x:Name="UpdateStatus" Text="Verification en cours..." Foreground="#C2D4CA" FontSize="12" TextWrapping="Wrap" Margin="0,0,0,8"/>
'     <StackPanel Orientation="Horizontal">
'      <Button x:Name="UpdateButton" Content="AUCUNE MISE A JOUR" IsEnabled="False"/>
'      <Button x:Name="UpdateLater" Content="PLUS TARD" Background="#1D2830" Foreground="#D1DEE0" BorderBrush="#48606C" Visibility="Collapsed"/>
'     </StackPanel>
'    </StackPanel>
'   </Border>
'   <Border Background="#0C1117" BorderBrush="#2B3935" BorderThickness="1" CornerRadius="7" Padding="14,10" Margin="0,0,0,12">
'    <StackPanel>
'     <TextBlock Foreground="#A5B9AF" FontFamily="Consolas" FontSize="13" FontWeight="Bold" Text="&gt; JOURNAL D'ANALYSE / EN DIRECT" Margin="0,0,0,6"/>
'     <TextBox x:Name="LogBox" IsReadOnly="True" Height="112" Background="#080C10" Foreground="#B6E5C6" BorderThickness="0"
'         FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Padding="4"/>
'    </StackPanel>
'   </Border>
'   <Border Background="#141B21" BorderBrush="#354654" BorderThickness="1" CornerRadius="7" Padding="15,11" Margin="0,0,0,10">
'    <Grid>
'     <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
'     <StackPanel>
'      <TextBlock Text="&gt; SURVEILLANCE AUTOMATIQUE" Foreground="#DFE9EA" FontFamily="Consolas" FontWeight="Bold" FontSize="13" Margin="0,0,0,8"/>
'      <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
'       <CheckBox x:Name="EnableMonitor" Content="ACTIVE" Margin="0,0,23,0" VerticalAlignment="Center"/>
'       <RadioButton x:Name="Weekly" GroupName="Frequence" Content="HEBDOMADAIRE" IsChecked="True" VerticalAlignment="Center"/>
'       <RadioButton x:Name="Daily" GroupName="Frequence" Content="QUOTIDIENNE" VerticalAlignment="Center"/>
'      </StackPanel>
'      <TextBlock x:Name="MonitorLabel" Text="Lecture du Planificateur..." Foreground="#99ADBC" FontSize="12" Margin="0,7,0,0"/>
'     </StackPanel>
'     <Button x:Name="SaveMonitor" Grid.Column="1" Content="ENREGISTRER" VerticalAlignment="Center" Margin="15,0,0,0"/>
'    </Grid>
'   </Border>
'   <TextBlock Text="CONTROLE SANS ELEVATION  /  AUCUNE ACTIVATION SILENCIEUSE  /  CLE COMPLETE JAMAIS AFFICHEE" FontFamily="Consolas"
'     Foreground="#768B82" FontSize="11" Margin="4,0,0,0"/>
'  </StackPanel>
' </ScrollViewer>
'</Window>
''@
'$reader = New-Object System.Xml.XmlNodeReader $xml
'$window = [System.Windows.Markup.XamlReader]::Load($reader)
'$names = @('TopStatus','TopTime','WinBadge','OfficeBadge','WinFields','OfficeFields','AdviceText','ActionButton','RefreshButton','ConfirmPanel','ConfirmText','YesButton','NoButton','LogBox','EnableMonitor','Weekly','Daily','MonitorLabel','SaveMonitor','UpdateStatus','UpdateButton','UpdateLater')
'$ui = @{}
'foreach ($n in $names) { $ui[$n] = $window.FindName($n) }
'$script:ScanProc = $null
'$script:ScanTimer = $null
'$script:RepairProc = $null
'$script:RepairTimer = $null
'$script:CloseTimer = $null
'$script:Report = $null
'$script:AdviceAction = 'None'
'$script:UpdateData = $null
'$script:UpdateProc = $null
'$script:UpdateTimer = $null
'$script:OnConfirm = $null
'$script:LogLines = New-Object 'System.Collections.Generic.List[string]'
'
'function ULog([string]$msg) {
'    $script:LogLines.Add($msg)
'    if ($script:LogLines.Count -gt 140) { $script:LogLines.RemoveAt(0) }
'    $ui.LogBox.Text = ($script:LogLines -join "`r`n")
'    $ui.LogBox.ScrollToEnd()
'}
'function Color-Of([string]$status) {
'    switch ($status) { 'ACTIVE' {'#8DF3AD'} 'INACTIVE' {'#FF8E88'} 'ABSENT' {'#FFB26C'} 'PARTIEL' {'#FFB26C'} 'ATTENTION' {'#F5C879'} 'INCONNU' {'#F5C879'} default {'#D4DFDE'} }
'}
'function New-InfoRow($box, [string]$label, [string]$value) {
'    $g = New-Object System.Windows.Controls.Grid
'    $g.Margin = [System.Windows.Thickness]::new(0,0,0,7)
'    $left = New-Object System.Windows.Controls.ColumnDefinition
'    $left.Width = [System.Windows.GridLength]::new(119)
'    $right = New-Object System.Windows.Controls.ColumnDefinition
'    $right.Width = [System.Windows.GridLength]::new(1,[System.Windows.GridUnitType]::Star)
'    [void]$g.ColumnDefinitions.Add($left); [void]$g.ColumnDefinitions.Add($right)
'    $l = New-Object System.Windows.Controls.TextBlock
'    $l.Text=$label; $l.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFrom('#829C93'); $l.FontSize=12
'    $r = New-Object System.Windows.Controls.TextBlock
'    $r.Text=if ($value) { $value } else { 'Non disponible' }
'    $r.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFrom('#E1ECE9')
'    $r.FontFamily=[System.Windows.Media.FontFamily]::new('Consolas')
'    $r.FontSize=12; $r.TextWrapping='Wrap'
'    [System.Windows.Controls.Grid]::SetColumn($r,1)
'    [void]$g.Children.Add($l); [void]$g.Children.Add($r); [void]$box.Children.Add($g)
'}
'function Set-Details($data) {
'    if (-not $data) { return }
'    $script:Report=$data
'    $ui.WinFields.Children.Clear(); $ui.OfficeFields.Children.Clear()
'    $w=$data.Windows; $o=$data.Office
'    $ui.WinBadge.Text='[ ' + $w.State + ' ]'; $ui.OfficeBadge.Text='[ ' + $o.State + ' ]'
'    $bc=New-Object System.Windows.Media.BrushConverter
'    $ui.WinBadge.Foreground=$bc.ConvertFrom((Color-Of $w.State))
'    $ui.OfficeBadge.Foreground=$bc.ConvertFrom((Color-Of $o.State))
'    New-InfoRow $ui.WinFields 'Edition' $w.Edition
'    New-InfoRow $ui.WinFields 'Version' $w.Build
'    New-InfoRow $ui.WinFields 'Activation' $w.State
'    New-InfoRow $ui.WinFields 'Type licence' $w.Channel
'    New-InfoRow $ui.WinFields 'Cle (5 fin)' $w.Key
'    New-InfoRow $ui.WinFields 'Expiration' $w.Expiry
'    New-InfoRow $ui.WinFields 'Derniere activ.' $w.Last
'    if ($w.Reason) { New-InfoRow $ui.WinFields 'Code erreur' $w.Reason }
'    New-InfoRow $ui.OfficeFields 'Edition' $o.Edition
'    New-InfoRow $ui.OfficeFields 'Version' $o.Version
'    New-InfoRow $ui.OfficeFields 'Activation' $o.State
'    New-InfoRow $ui.OfficeFields 'Type licence' $o.Type
'    New-InfoRow $ui.OfficeFields 'Cle (5 fin)' $o.Key
'    New-InfoRow $ui.OfficeFields 'Expiration' $o.Expiry
'    New-InfoRow $ui.OfficeFields 'Derniere activ.' $o.Last
'    New-InfoRow $ui.OfficeFields 'Applications' $o.Apps
'    $script:AdviceAction = [string]$data.Advice.Action
'    $ui.AdviceText.Text = [string]$data.Advice.Message
'    $ui.ActionButton.Content = [string]$data.Advice.Button
'    $ui.ActionButton.IsEnabled = ($script:AdviceAction -ne 'None')
'    $ui.TopStatus.Text=if ($data.Advice.Alert) {'[ INTERVENTION CONSEILLEE ]'} else {'[ CONTROLE TERMINE ]'}
'    $statusColor = if ($data.Advice.Alert) { '#F5C879' } else { '#8DF3AD' }
'    $ui.TopStatus.Foreground=$bc.ConvertFrom($statusColor)
'    try { $ui.TopTime.Text='Dernier controle : '+([datetime]$data.Timestamp).ToString('dd/MM/yyyy HH:mm:ss') } catch { }
'    $script:UpdateData = $data.Update
'    if ($data.Update -and $data.Update.Available -and $data.Update.Verified) {
'        $ui.UpdateStatus.Text = ('Version installee : {0}  |  Nouvelle : {1}  |  {2}' -f $AppVersion,$data.Update.Version,$data.Update.Notes)
'        $ui.UpdateButton.Content = ('INSTALLER LA V' + $data.Update.Version)
'        $ui.UpdateButton.IsEnabled = $true
'        $ui.UpdateLater.Visibility='Visible'
'    } elseif ($data.Update -and $data.Update.Verified) {
'        $ui.UpdateStatus.Text = ('Version {0} : a jour. Signature verifiee.' -f $AppVersion)
'        $ui.UpdateButton.Content='AUCUNE MISE A JOUR'
'        $ui.UpdateButton.IsEnabled=$false
'        $ui.UpdateLater.Visibility='Collapsed'
'    } else {
'        $ui.UpdateStatus.Text=if ($data.Update) { $data.Update.Status } else { 'Version inconnue; reessayez le controle.' }
'        $ui.UpdateButton.Content='MAJ INDISPONIBLE'
'        $ui.UpdateButton.IsEnabled=$false
'        $ui.UpdateLater.Visibility='Collapsed'
'    }
'}
'function Hide-Confirm { $ui.ConfirmPanel.Visibility='Collapsed'; $script:OnConfirm=$null }
'function Ask-InWindow([string]$question, [scriptblock]$approved) {
'    $script:OnConfirm=$approved
'    $ui.ConfirmText.Text=$question
'    $ui.ConfirmPanel.Visibility='Visible'
'    $ui.YesButton.Focus() | Out-Null
'}
'function Confirm-Yes {
'    $callback=$script:OnConfirm
'    Hide-Confirm
'    if ($callback) { & $callback }
'}
'function Confirm-No { Hide-Confirm; ULog '[USER] Intervention annulee. Aucune modification.' }
'$ui.YesButton.Add_Click({ Confirm-Yes })
'$ui.NoButton.Add_Click({ Confirm-No })
'$window.Add_PreviewKeyDown({
'    param($s,$e)
'    if ($ui.ConfirmPanel.Visibility -ne 'Visible') { return }
'    if ($e.Key -eq [System.Windows.Input.Key]::D1 -or $e.Key -eq [System.Windows.Input.Key]::NumPad1) { Confirm-Yes; $e.Handled=$true }
'    elseif ($e.Key -eq [System.Windows.Input.Key]::D0 -or $e.Key -eq [System.Windows.Input.Key]::NumPad0) { Confirm-No; $e.Handled=$true }
'})
'function Refresh-MonitorLabel {
'    $m=Get-Monitor
'    $ui.EnableMonitor.IsChecked=[bool]$m.Enabled
'    $ui.Daily.IsChecked=($m.Frequency -eq 'daily')
'    $ui.Weekly.IsChecked=($m.Frequency -ne 'daily')
'    $ui.MonitorLabel.Text=('Etat actuel : {0}  |  12 h, session utilisateur ouverte' -f $m.Label)
'}
'$ui.SaveMonitor.Add_Click({
'    try {
'        $enabled=[bool]$ui.EnableMonitor.IsChecked
'        $freq=if ($ui.Daily.IsChecked) { 'daily' } else { 'weekly' }
'        Set-Monitor $enabled $freq
'        Refresh-MonitorLabel
'        ULog ('[OK] Surveillance ' + $(if ($enabled) {$freq} else {'desactivee'}) + ' - appliquee.')
'    } catch { ULog ('[ERREUR] Surveillance : ' + $_.Exception.Message) }
'})
'function Start-Scan {
'    if ($script:ScanProc -and -not $script:ScanProc.HasExited) { ULog '[INFO] Une analyse est deja en cours.'; return }
'    Hide-Confirm
'    $ui.TopStatus.Text='[ ANALYSE EN COURS ]'
'    $ui.TopStatus.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFrom('#F5C879')
'    $ui.ActionButton.IsEnabled=$false
'    $ui.ActionButton.Content='ANALYSE EN COURS...'
'    $ui.RefreshButton.IsEnabled=$false
'    $id=[guid]::NewGuid().ToString('N')
'    $report=Join-Path $Root ('scan-' + $id + '.json')
'    $progress=Join-Path $Root ('scan-' + $id + '.log')
'    $script:LogLines.Clear()
'    ULog '[START] Inspection des licences Windows / Office...'
'    $start=New-Object System.Diagnostics.ProcessStartInfo
'    $start.FileName=$PSExe
'    $start.Arguments=('-NoLogo -NoProfile -NonInteractive -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Mode Scan -ReportFile "{1}" -ProgressFile "{2}"' -f $PSCommandPath,$report,$progress)
'    $start.UseShellExecute=$false; $start.CreateNoWindow=$true; $start.WindowStyle=[System.Diagnostics.ProcessWindowStyle]::Hidden
'    try {
'        $script:ScanProc=[System.Diagnostics.Process]::Start($start)
'    } catch { ULog ('[ERREUR] Impossible de lancer le scan : ' + $_.Exception.Message); $ui.RefreshButton.IsEnabled=$true; return }
'    $script:LastProgress=''
'    $script:ScanReportFile=$report
'    $script:ScanProgressFile=$progress
'    if ($script:ScanTimer) { $script:ScanTimer.Stop() }
'    $script:ScanTimer=New-Object System.Windows.Threading.DispatcherTimer
'    $script:ScanTimer.Interval=[timespan]::FromMilliseconds(400)
'    $script:ScanTimer.Add_Tick({
'        if (Test-Path -LiteralPath $script:ScanProgressFile) {
'            try {
'                $raw=Get-Content -LiteralPath $script:ScanProgressFile -Raw -Encoding UTF8
'                if ($raw -ne $script:LastProgress) {
'                    $script:LastProgress=$raw
'                    $ui.LogBox.Text="[START] Inspection des licences Windows / Office...`r`n" + $raw
'                    $ui.LogBox.ScrollToEnd()
'                }
'            } catch { }
'        }
'        if ($script:ScanProc.HasExited) {
'            $script:ScanTimer.Stop()
'            $ui.RefreshButton.IsEnabled=$true
'            $scanResult=Read-Json $script:ScanReportFile
'            if ($scanResult) {
'                Set-Details $scanResult
'                # The full process log is retained in the same window.
'            } else {
'                $ui.TopStatus.Text='[ ERREUR DE CONTROLE ]'
'                $ui.AdviceText.Text='Diagnostic incomplet. Consultez le journal; aucun changement de licence effectue.'
'                ULog '[ERREUR] Rapport de diagnostic non produit.'
'            }
'            try { Remove-Item -LiteralPath $script:ScanReportFile,$script:ScanProgressFile -Force -ErrorAction SilentlyContinue } catch { }
'            if ($script:LastProgress) {
'                foreach ($line in ($script:LastProgress -split '[\r\n]+')) {
'                    if ($line.Trim()) { $script:LogLines.Add($line.Trim()) }
'                }
'            }
'            ULog '[DONE] Diagnostic termine; tableau actualise.'
'            $script:ScanProc.Dispose(); $script:ScanProc=$null
'        }
'    })
'    $script:ScanTimer.Start()
'}
'$ui.RefreshButton.Add_Click({ Start-Scan })
'$ui.UpdateLater.Add_Click({
'    if (-not ($script:UpdateData -and $script:UpdateData.Available)) { return }
'    Save-Json $UpdateSnooze ([pscustomobject]@{ Version=$script:UpdateData.Version; Timestamp=(Get-Date).ToString('o') })
'    ULog ('[MAJ] Rappel de la V' + $script:UpdateData.Version + ' reporte de 7 jours.')
'    $ui.UpdateButton.IsEnabled=$false
'    $ui.UpdateLater.Visibility='Collapsed'
'    $ui.UpdateStatus.Text='Mise a jour reportee a une verification ulterieure. Version actuelle conservee.'
'})
'$ui.UpdateButton.Add_Click({
'    if (-not ($script:UpdateData -and $script:UpdateData.Available -and $script:UpdateData.Verified)) { return }
'    $targetVersion=[string]$script:UpdateData.Version
'    Ask-InWindow ('MISE A JOUR V' + $targetVersion + ' : telecharger et installer la nouvelle version signee ? Aucun changement sans votre confirmation. La fenetre sera relancee apres verification.') {
'        try {
'            if (Test-Path -LiteralPath $UpdateResult) { Remove-Item -LiteralPath $UpdateResult -Force }
'            $ui.UpdateButton.IsEnabled=$false
'            $ui.UpdateLater.IsEnabled=$false
'            ULog '[MAJ] Accord donne. Verification de la signature et telechargement dans un processus invisible...'
'            $start=New-Object System.Diagnostics.ProcessStartInfo
'            $start.FileName=$PSExe
'            $start.Arguments=('-NoLogo -NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -Mode InstallUpdate' -f $PSCommandPath)
'            $start.UseShellExecute=$false
'            $start.CreateNoWindow=$true
'            $start.WindowStyle=[System.Diagnostics.ProcessWindowStyle]::Hidden
'            $script:UpdateProc=[System.Diagnostics.Process]::Start($start)
'            $script:UpdateTimer=New-Object System.Windows.Threading.DispatcherTimer
'            $script:UpdateTimer.Interval=[timespan]::FromMilliseconds(800)
'            $script:UpdateTimer.Add_Tick({
'                if ($script:UpdateProc -and -not $script:UpdateProc.HasExited) { return }
'                $script:UpdateTimer.Stop()
'                $result=Read-Json $UpdateResult
'                if ($result -and $result.Success) {
'                    ULog ('[MAJ] ' + $result.Message)
'                    try {
'                        Start-Process -FilePath (Join-Path $env:WINDIR 'System32\wscript.exe') -ArgumentList ('"{0}"' -f $InstalledEntry)
'                        $window.Close()
'                    } catch { ULog ('[ERREUR] Redemarrage : ' + $_.Exception.Message); $ui.UpdateButton.IsEnabled=$true }
'                } else {
'                    ULog ('[ERREUR] ' + $(if ($result) { $result.Message } else { 'Aucun rapport de mise a jour.' }))
'                    $ui.UpdateButton.IsEnabled=$true
'                    $ui.UpdateLater.IsEnabled=$true
'                }
'                if ($script:UpdateProc) { $script:UpdateProc.Dispose(); $script:UpdateProc=$null }
'            })
'            $script:UpdateTimer.Start()
'        } catch { ULog ('[ERREUR] Preparation de la mise a jour : ' + $_.Exception.Message); $ui.UpdateButton.IsEnabled=$true }
'    }
'})
'function Office-StillOpen {
'    return (@(Get-OfficeProcesses).Count -gt 0)
'}
'function Start-Repair([string]$what) {
'    $script:PendingTarget=$what
'    Ask-InWindow ('CONFIRMATION : tenter une reactivation officielle de ' + $what + ' avec la licence existante ? Windows affichera une demande administrateur (UAC). Aucune cle ne sera ajoutee, aucun outil MAS ne sera execute.') {
'        try {
'            if (Test-Path -LiteralPath $RepairResult) { Remove-Item -LiteralPath $RepairResult -Force }
'            $args=('"{0}" /repair {1}' -f $InstalledEntry,$script:PendingTarget)
'            $script:RepairProc=Start-Process -FilePath (Join-Path $env:WINDIR 'System32\wscript.exe') -ArgumentList $args -Verb RunAs -PassThru
'            ULog ('[ACTION] Demande UAC confirmee : activation ' + $script:PendingTarget + '...')
'            $ui.ActionButton.IsEnabled=$false
'            $script:RepairTimer=New-Object System.Windows.Threading.DispatcherTimer
'            $script:RepairTimer.Interval=[timespan]::FromSeconds(1)
'            $script:RepairTimer.Add_Tick({
'                if (-not $script:RepairProc -or -not $script:RepairProc.HasExited) { return }
'                $script:RepairTimer.Stop()
'                $result=Read-Json $RepairResult
'                if ($result) { ULog ('[RESULTAT] ' + $result.Message) }
'                else { ULog '[INFO] Aucune confirmation du processus eleve. Intervention non verifiee.' }
'                $script:RepairProc.Dispose(); $script:RepairProc=$null
'                Start-Scan
'            })
'            $script:RepairTimer.Start()
'        } catch { ULog ('[ERREUR] Elevation refusee ou indisponible : ' + $_.Exception.Message); $ui.ActionButton.IsEnabled=$true }
'    }
'}
'function Check-Office-ThenRepair([string]$what) {
'    $script:PendingTarget=$what
'    $open=@(Get-OfficeProcesses)
'    if ($open.Count -eq 0) { Start-Repair $what; return }
'    $names=($open | Select-Object -ExpandProperty ProcessName -Unique) -join ', '
'    Ask-InWindow ('Applications Office ouvertes : ' + $names + '. Enregistrez vos documents. Demander leur fermeture normale maintenant ? Aucune fermeture forcee ne sera effectuee.') {
'        ULog '[ACTION] Fermeture normale demandee. Sauvegardez les documents dans Office.'
'        foreach ($p in @(Get-OfficeProcesses)) { try { [void]$p.CloseMainWindow() } catch { } }
'        $script:OfficeCloseTickCounter=0
'        if ($script:CloseTimer) { $script:CloseTimer.Stop() }
'        $script:CloseTimer=New-Object System.Windows.Threading.DispatcherTimer
'        $script:CloseTimer.Interval=[timespan]::FromSeconds(1)
'        $script:CloseTimer.Add_Tick({
'            $script:OfficeCloseTickCounter++
'            if (-not (Office-StillOpen)) { $script:CloseTimer.Stop(); ULog '[OK] Office est ferme.'; Start-Repair $script:PendingTarget }
'            elseif ($script:OfficeCloseTickCounter -ge 90) { $script:CloseTimer.Stop(); ULog '[ATTENTION] Office est toujours ouvert. Fermez-le puis relancez l''action.' }
'        })
'        $script:CloseTimer.Start()
'    }
'}
'$ui.ActionButton.Add_Click({
'    switch ($script:AdviceAction) {
'        'WindowsRepair' { Check-Office-ThenRepair 'Windows' }
'        'OfficeRepair' { Check-Office-ThenRepair 'Office' }
'        'OfficeInstall' {
'            Ask-InWindow 'Ouvrir le portail Microsoft pour installer la suite liee a votre licence ? Vous choisirez votre edition; aucun installateur ne sera execute silencieusement.' {
'                ULog '[ACTION] Ouverture du portail officiel Microsoft.'
'                Start-Process 'https://www.microsoft.com/microsoft-365/download-office'
'            }
'        }
'        'OfficeAccount' {
'            Ask-InWindow 'Ouvrir Word pour verifier Fichier > Compte et reconnecter un compte Microsoft autorise ?' {
'                $loc=@(Office-Roots) | ForEach-Object { Join-Path $_ 'WINWORD.EXE' } | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
'                if ($loc) { Start-Process -FilePath $loc; ULog '[ACTION] Word ouvert : Fichier > Compte.' }
'                else { ULog '[ATTENTION] Word introuvable. Utilisez le portail Microsoft.' }
'            }
'        }
'        'WindowsSettings' {
'            Ask-InWindow 'Ouvrir les parametres Windows > Systeme > Activation ?' { Start-Process 'ms-settings:activation' }
'        }
'    }
'})
'$window.Add_ContentRendered({
'    try {
'        $previous=Get-Monitor
'        if ($previous.Enabled -and $previous.Legacy) {
'            Set-Monitor $true $previous.Frequency
'            ULog '[MIGRATION] Ancienne tache planifiee convertie vers le moniteur V3.'
'        }
'    } catch { ULog ('[ATTENTION] Migration de la surveillance : ' + $_.Exception.Message) }
'    Refresh-MonitorLabel
'    if ($UseReport) {
'        $prior=Read-Json $LastReport
'        if ($prior) { Set-Details $prior; ULog '[MONITEUR] Une anomalie de licence a ete detectee.'; ULog '[INFO] Le resultat est celui du controle planifie.'; return }
'    }
'    Start-Scan
'})
'$window.Add_Closed({
'    try { if ($script:ScanTimer) { $script:ScanTimer.Stop() }; if ($script:ScanProc -and -not $script:ScanProc.HasExited) { $script:ScanProc.Kill() } } catch { }
'})
'[void]$window.ShowDialog()
'} catch {
'    Add-History ('ERREUR interface V3 : ' + $_.Exception.Message)
'    try { [System.Windows.MessageBox]::Show(('Interface impossible : ' + $_.Exception.Message), 'Moniteur de licences') | Out-Null } catch { }
'}
'@@END_POWERSHELL
