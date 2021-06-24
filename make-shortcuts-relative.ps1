Param([switch] $deleteShortcuts = $False, [switch] $dryRun = $False, [switch] $recursive)

$current = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object System.Security.Principal.WindowsPrincipal($current)
if (-not $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
	$pinfo = New-Object System.Diagnostics.ProcessStartInfo
	$pinfo.ArgumentList.Add('-File')
	$pinfo.ArgumentList.Add($PSCommandPath)
	if ($deleteShortcuts) {
		$pinfo.ArgumentList.Add('-deleteShortcuts')
	}
	if ($dryRun) {
		$pinfo.ArgumentList.Add('-dryRun')
	}
	if ($recursive) {
		$pinfo.ArgumentList.Add('-recursive')
	}
	$pinfo.FileName = 'pwsh.exe'
	$pinfo.UseShellExecute = $True
	$pinfo.Verb = 'RunAs'
	$pinfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
	$pinfo.WorkingDirectory = Get-Location
	$process = New-Object System.Diagnostics.Process
	$process.StartInfo = $pinfo
	$result = $process.Start()
	$process.WaitForExit()
	exit $result
}

$children = $recursive ? (Get-ChildItem -Recurse .) : (Get-ChildItem .)

$cwd = Get-Location
$cwd = $cwd.Path.Split('\')
$shell = New-Object -ComObject WScript.Shell

$toBeRemoved = @()
$children | Where-Object { -not $_.PsIsContainer -and $_.FullName.EndsWith('.lnk') } | ForEach-Object {
	$link = $shell.CreateShortcut($_.FullName)
	if (-not $link.TargetPath) {
		Clear-Variable link
		Write-Output "${_.FullName} has no target path, so it has been omitted."
		return
	}
	$from = $link.FullName.Split('\')
	$path = @()
	for ($i = $cwd.count; $i -lt $from.count - 1; $i++) {
		$path += $from[$i]
	}
	$to = $link.TargetPath.Split('\')
	if ($to[0] -ne $cwd[0]) {
		$drive = '"' + $to[0] + '"'
		Write-Output "${link.FullName} is linked to other drive ${drive}, so it has been omitted."
		return
	}
	$relative = @()
	for ($i = 0; $i -lt [Math]::Min($from.count, $to.count); $i++) {
		if ($from[$i] -ne $to[$i]) {
			for ($j = $i; $j -lt $from.count - 1; $j++) {
				$relative += '..'
			}
			for ($j = $i; $j -lt $to.count; $j++) {
				$relative += $to[$j]
			}
			break
		}
	}
	$relative = $relative -join '\'
	$path += $from[$from.count - 1] -creplace '.lnk', ''
	$path = $path -join '\'
	$item = Get-Item -ErrorAction SilentlyContinue $path
	if ($item) {
		Clear-Variable item
		Write-Output "${path} already exists, so it has been omitted."
		return
	}
	Clear-Variable item
	$target = Get-Item $link.TargetPath
	if ($target.PsIsContainer) {
		Clear-Variable target
		$path = '"' + $path + '"'
		$relative = '"' + $relative + '"'
		if (-not $dryRun) {
			cmd /C "mklink /D ${path} ${relative}"
		}
		$toBeRemoved += $_.FullName
	}
	else {
		Clear-Variable target
		if (-not $dryRun) {
			New-Item -ItemType SymbolicLink -Name $path -Value $relative
		}
		$toBeRemoved += $_.FullName
	}
}

Clear-Variable children
Clear-Variable cwd
Clear-Variable shell

if ($deleteShortcuts -and -not $dryRun) {
	$toBeRemoved | ForEach-Object {
		Remove-Item $_ -ErrorAction SilentlyContinue
	}
}

Clear-Variable toBeRemoved
