param (
    $Replaces = @(
        @('^Alpine', 'Alpine'),
        @('^Alpine-', 'Alpine '),
        @('^Ubuntu', 'Ubuntu'),
        @('^Ubuntu(\d{2})(\d{2})$', 'Ubuntu $1.$2')
    ),
    [switch] $Delete
)
Set-Location $PSScriptRoot

# REGEX (to be applied to the `BaseName`s of *.exe files)
# See: https://vexx32.github.io/2020/02/15/Building-Arrays-Collections/#using-the-pipeline
$regex = $(foreach ($r in $Replaces) {$r[0]}) -join '|'

# TARGET
$file = Get-ChildItem *.exe `
        | Sort-Object LastWriteTime -Descending `
        | Where-Object {$_.BaseName -match $regex} `
        | Select-Object -First 1
if (-not $file) {
    $msg = "EXE file is not found"
    $info = "Dir:    `"$PSScriptRoot`"`nRegex:  `"$regex`""
    [Console]::Error.WriteLine("$msg`n$info")
    exit 1
}
$target = $file.FullName

# LINK
$name = $file.BaseName
foreach ($r in $Replaces) {
    $name = $name -replace $r
}
# See: https://stackoverflow.com/questions/31747115/powershell-desktop-variable
# Actual value: "%APPDATA%\Microsoft\Windows\Start Menu" (from Vista to 11)
$startmenu = [Environment]::GetFolderPath('StartMenu')
$link = "$startmenu\Programs\$name.lnk"

if ($Delete -and -not (Test-Path $link)) {
    $msg = "Shortcut does not exist in the Start menu"
    $info = "Link:   `"$link`""
    [Console]::Error.WriteLine("$msg`n$info")
    exit 1
}

# WScript.Shell INTEROP
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($link)

if ($Delete) {
    # DELETE shortcut
    $target = $shortcut.TargetPath  # Get REAL TARGET of LINK
    $sha1 = (Get-FileHash $link -Algorithm SHA1).Hash.ToLower()
    Remove-Item $link
    if ($?) {
        $msg = "Shortcut has been deleted from the Start menu"
        $ret = 0
    } else {
        $msg = "An error occurred while deleting the shortcut from the Start menu"
        $ret = 1
    }
    $info = "Link:   `"$link`"`n        ($sha1)`nTarget: `"$target`""
} else {
    # CREATE shortcut
    $shortcut.TargetPath = $target
    $shortcut.IconLocation = "$target,0"
    $shortcut.Save()
    if ($?) {
        $msg = "Shortcut has been created in the Start menu"
        $ret = 0
        $sha1 = (Get-FileHash $link -Algorithm SHA1).Hash.ToLower()
    } else {
        $msg = "An error occurred while creating the shortcut in the Start menu"
        $ret = 1
    }
    $info = "Target: `"$target`"`nLink:   `"$link`""
    if ($sha1) {$info += "`n        ($sha1)"}
}

if ($ret -eq 0) {
    "$info`n$msg"
} else {
    [Console]::Error.WriteLine("$msg`n$info")
}
exit $ret
