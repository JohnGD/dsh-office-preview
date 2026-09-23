# Installs dsh-office-preview (Word .docx / Excel .xlsx sidebar preview) into a
# DSH Desktop / Harness home.
#
#   powershell -ExecutionPolicy Bypass -File install.ps1
#   powershell -ExecutionPolicy Bypass -File install.ps1 -HarnessHome "D:\dsh-home" -Profile web
#
# Three steps, nothing else:
#   1. copy the package to <home>\plugins\dsh-office-preview
#   2. link <home>\profiles\node_modules\dsh-office-preview at it, which is a
#      root the Cordis loader resolves plugin entry names through
#   3. add one `insert` row to the profile's own patch layer
#      (<home>\profiles\<profile>\cordis.patch.yml), backing that file up first
#
# The patch layer is watched (`patchReload: live`), so the plugin usually mounts
# without a restart; refresh the Harness page to load the client bundle.
param(
	[string]$HarnessHome,
	[string]$Profile
)

$ErrorActionPreference = 'Stop'
$PluginName = 'dsh-office-preview'
$source = $PSScriptRoot

# ---- 1. locate the Harness home -------------------------------------------------
$candidates = @()
if ($HarnessHome) { $candidates += $HarnessHome }
if ($env:DSH_HOME) { $candidates += $env:DSH_HOME }
$candidates += (Join-Path $env:APPDATA 'dsh-desktop\harness')
$candidates += (Join-Path $env:USERPROFILE '.dsh')
$candidates += (Join-Path $env:USERPROFILE 'Library\Application Support\dsh-desktop\harness')

$dshRoot = $null
foreach ($candidate in $candidates) {
	if ($candidate -and (Test-Path (Join-Path $candidate 'profiles'))) { $dshRoot = (Resolve-Path $candidate).Path; break }
}
if (-not $dshRoot) { throw "Harness home not found. Looked at: $($candidates -join '; '). Pass -HarnessHome <path>." }
Write-Host "harness home : $dshRoot"

# ---- 2. locate the profile ------------------------------------------------------
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
$patchPath = Join-Path $profilesRoot "$profileName\cordis.patch.yml"
Write-Host "profile      : $profileName"

# ---- 3. copy the package --------------------------------------------------------
$target = Join-Path $dshRoot "plugins\$PluginName"
New-Item -ItemType Directory -Force -Path $target, (Join-Path $target 'lib') | Out-Null
foreach ($entry in @('package.json', 'cordis.patch.yml', 'dsh.plugin.json')) {
	$from = Join-Path $source $entry
	if (Test-Path $from) { Copy-Item $from (Join-Path $target $entry) -Force }
}
# Every markdown file, whatever it is called in the reader's language.
Get-ChildItem -Path (Join-Path $source '*.md') | Copy-Item -Destination $target -Force
Copy-Item (Join-Path $source 'lib\*.js') (Join-Path $target 'lib') -Force
Write-Host "files        : $target"

# ---- 4. link it into the loader's resolution root -------------------------------
$link = Join-Path $profilesRoot "node_modules\$PluginName"
New-Item -ItemType Directory -Force -Path (Split-Path $link) | Out-Null
if (Test-Path $link) {
	$item = Get-Item $link -Force
	if ($item.LinkType) { $item.Delete() } else { Remove-Item $link -Recurse -Force }
}
if ($env:OS -eq 'Windows_NT') {
	New-Item -ItemType Junction -Path $link -Target $target | Out-Null
} else {
	New-Item -ItemType SymbolicLink -Path $link -Target $target | Out-Null
}
Write-Host "link         : $link -> $target"

# ---- 5. mount it through the profile's own patch layer --------------------------
# A profile that already declares this package as a dependency mounts it as a
# bundle layer instead. Adding the patch row there too would mount one entry id
# twice, so the row is only written for the dependency-free form.
$manifestText = Get-Content (Join-Path $profilesRoot "$profileName\package.json") -Raw
if ($manifestText -match ('"' + [regex]::Escape($PluginName) + '"\s*:')) {
	Write-Host "patch layer  : skipped - $PluginName is already a profile dependency (market-managed); its own bundle patch mounts it"
} else {
	$patch = Get-Content $patchPath -Raw
	if ($patch -match [regex]::Escape($PluginName)) {
		Write-Host "patch layer  : already names $PluginName; left unchanged"
	} else {
		$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
		Copy-Item $patchPath "$patchPath.bak-$stamp" -Force
		$body = [regex]::Replace($patch, '(?m)^\s*\[\s*\]\s*$', '').TrimEnd()
		$row = "# Office preview (Word .docx / Excel .xlsx) for the sidebar document pane.`n- insert:`n    - id: $PluginName`n      name: $PluginName`n"
		$next = if ($body -eq '') { $row } else { "$body`n`n$row" }
		# UTF-8 without a byte-order mark: the Harness reads this YAML as UTF-8.
		[System.IO.File]::WriteAllText($patchPath, $next, [System.Text.UTF8Encoding]::new($false))
		Write-Host "patch layer  : updated $patchPath (backup: $patchPath.bak-$stamp)"
	}
}

Write-Host ''
Write-Host 'done. Refresh the Harness page (Ctrl+R). If the renderer does not appear, restart DSH Desktop.'
Write-Host "self-check: node `"$source\test\parse.mjs`" `"$source\samples\sample.xlsx`" `"$source\samples\sample.docx`""
