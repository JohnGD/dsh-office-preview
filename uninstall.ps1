# Removes dsh-office-preview from a DSH Desktop / Harness home.
#
#   powershell -ExecutionPolicy Bypass -File uninstall.ps1
#   powershell -ExecutionPolicy Bypass -File uninstall.ps1 -RemoveFiles
param(
	[string]$HarnessHome,
	[string]$Profile,
	[switch]$RemoveFiles
)

$ErrorActionPreference = 'Stop'
$PluginName = 'dsh-office-preview'

$candidates = @()
if ($HarnessHome) { $candidates += $HarnessHome }
if ($env:DSH_HOME) { $candidates += $env:DSH_HOME }
$candidates += (Join-Path $env:APPDATA 'dsh-desktop\harness')
$candidates += (Join-Path $env:USERPROFILE '.dsh')

$dshRoot = $null
foreach ($candidate in $candidates) {
	if ($candidate -and (Test-Path (Join-Path $candidate 'profiles'))) { $dshRoot = (Resolve-Path $candidate).Path; break }
}
if (-not $dshRoot) { throw "Harness home not found. Pass -HarnessHome <path>." }

$profilesRoot = Join-Path $dshRoot 'profiles'
$profileName = $Profile
if (-not $profileName) {
	if (Test-Path (Join-Path $profilesRoot 'web\cordis.patch.yml')) { $profileName = 'web' }
	else {
		$found = Get-ChildItem $profilesRoot -Directory -ErrorAction SilentlyContinue |
			Where-Object { Test-Path (Join-Path $_.FullName 'cordis.patch.yml') } |
			Select-Object -First 1
		if ($found) { $profileName = $found.Name }
	}
}
if (-not $profileName) { throw "No dsh profile with a cordis.patch.yml under $profilesRoot." }

$target = Join-Path $dshRoot "plugins\$PluginName"
$link = Join-Path $profilesRoot "node_modules\$PluginName"
$patchPath = Join-Path $profilesRoot "$profileName\cordis.patch.yml"

if (Test-Path $link) {
	$item = Get-Item $link -Force
	if ($item.LinkType) { $item.Delete() } else { Remove-Item $link -Recurse -Force }
	Write-Host "unlinked     : $link"
}

if (Test-Path $patchPath) {
	$patch = Get-Content $patchPath -Raw
	if ($patch -match [regex]::Escape($PluginName)) {
		$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
		Copy-Item $patchPath "$patchPath.bak-$stamp" -Force
		$name = [regex]::Escape($PluginName)
		# Remove the row this package's installer wrote, together with the comment
		# line it wrote above that row (the comment is optional for older installs).
		$next = [regex]::Replace($patch, "(?ms)^(?:#\s*Office preview[^\r\n]*\r?\n)?^[ \t]*- insert:[ \t]*\r?\n[ \t]*- id: $name[ \t]*\r?\n[ \t]*name: $name[ \t]*\r?\n?", '').TrimEnd()
		# A patch layer that is only comments parses as null rather than as an
		# array, so restore the empty-array placeholder when no entry is left.
		$entries = @($next -split "\r?\n" | Where-Object { $_ -notmatch '^\s*#' -and $_.Trim() -ne '' })
		if ($entries.Count -eq 0) { $next = if ($next -eq '') { '[]' } else { "$next`n[]" } }
		[System.IO.File]::WriteAllText($patchPath, "$next`n", [System.Text.UTF8Encoding]::new($false))
		Write-Host "patch layer  : cleaned $patchPath (backup: $patchPath.bak-$stamp)"
	}
}

if ($RemoveFiles -and (Test-Path $target)) {
	Remove-Item $target -Recurse -Force
	Write-Host "removed      : $target"
}

Write-Host ''
Write-Host 'done. Refresh the Harness page; if the renderer is still listed, restart DSH Desktop.'
