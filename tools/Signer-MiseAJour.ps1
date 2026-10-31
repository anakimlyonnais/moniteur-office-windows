# Signer-MiseAJour.ps1 -- utilisation hors ligne sur le PC de l'editeur.
# Ne JAMAIS ajouter la cle privee au depot public.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Fichier,
    [Parameter(Mandatory)][string]$Version,
    [Parameter(Mandatory)][string]$Commit,
    [Parameter(Mandatory)][string]$ClePriveeXml,
    [string]$Notes = 'Ameliorations et corrections',
    [string]$Sortie = 'latest.json'
)
$ErrorActionPreference = 'Stop'
if ($Version -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}$') { throw 'Version attendue : X.Y.Z' }
if ($Commit -notmatch '^[a-fA-F0-9]{40}$') { throw 'SHA GitHub attendu : 40 caracteres hexadecimaux' }
if ($Notes.Length -gt 1000) { throw 'Notes trop longues' }
if (-not (Test-Path -LiteralPath $Fichier)) { throw 'Fichier VBS absent' }
if (-not (Test-Path -LiteralPath $ClePriveeXml)) { throw 'Cle privee absente' }
$url = "https://raw.githubusercontent.com/anakimlyonnais/moniteur-office-windows/$Commit/Moniteur-Office-Windows.vbs"
$hash = (Get-FileHash -LiteralPath $Fichier -Algorithm SHA256).Hash.ToLowerInvariant()
$payload = "$Version`n$url`n$hash`n$Notes`n"
$rsa = New-Object System.Security.Cryptography.RSACryptoServiceProvider
try {
    $rsa.FromXmlString([IO.File]::ReadAllText((Resolve-Path -LiteralPath $ClePriveeXml)))
    $sig = [Convert]::ToBase64String($rsa.SignData([Text.Encoding]::UTF8.GetBytes($payload), 'SHA256'))
} finally { $rsa.Dispose() }
$manifest = [ordered]@{ version=$Version; url=$url; sha256=$hash; notes=$Notes; signature=$sig }
$json = $manifest | ConvertTo-Json -Depth 3
[IO.File]::WriteAllText([IO.Path]::GetFullPath($Sortie),$json+"\n",(New-Object Text.UTF8Encoding($false)))
Write-Host "Manifeste signe : $Sortie"
Write-Host "SHA-256 : $hash"
Write-Host "Publier updates/latest.json en DERNIER, apres la version VBS sur GitHub."
