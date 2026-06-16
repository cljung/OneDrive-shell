Import-Module Microsoft.Graph.Files

# Clear the screen to give a clean terminal experience
Clear-Host

$cwd = "root:"  # name of root folder
$DriveId = ""   # unique value for your OneDrive, retrieved at login

function print-help {
    Write-Host "Available Commands:" -ForegroundColor Cyan
    Write-Host "  login                          - Sign-in to OneDrive personal" -ForegroundColor Yellow
    Write-Host "  logoff                         - Sign-out from OneDrive personal" -ForegroundColor Yellow
    Write-Host "  cd <path>                      - Change current directory" -ForegroundColor Yellow
    Write-Host "  md/mkdir <name>                - Create a new directory in OneDrive. No wildcards supported" -ForegroundColor Yellow
    Write-Host "  rd/rmdir <name>                - Remove a directory from OneDrive. No wildcards supported" -ForegroundColor Yellow
    Write-Host "  dir/ls (<path>)                - List contents of current or named directory. No wildcards supported" -ForegroundColor Yellow
    Write-Host "  ul/upload <src> (<dst>) (-f)   - Upload file/folder to OneDrive. No wildcards supported" -ForegroundColor Yellow
    Write-Host "  dl/download <src> (<dst>) (-f) - Download file/folder from OneDrive. No wildcards supported" -ForegroundColor Yellow
    Write-Host "  del/rm <src> (-f)              - Delete file from OneDrive. No wildcards supported" -ForegroundColor Yellow
    Write-Host "  help                           - Show this help message" -ForegroundColor Yellow
    Write-Host "  exit                           - Exit the custom command prompt" -ForegroundColor Yellow
    Write-Host "  !command                       - Execute a local command" -ForegroundColor Yellow
    write-host "`nNotes:" -ForegroundColor Cyan
    write-host "  - Paths are relative to current directory unless starting with '/'" -ForegroundColor Yellow
    write-host "  - No wildcards supported in OneDrive paths" -ForegroundColor Yellow
    write-host "  - root:/ is prepended automatically for OneDrive paths" -ForegroundColor Yellow
    write-host "  - -f/--force                   - Force operation without confirmation, overwriting existing files" -ForegroundColor Yellow
}

function Create-OneDriveFolder {
    param (
        [string]$DriveId,
        [string]$Path
    )
    if (!$Path.StartsWith("root:")) {
        $Path = "root:/$Path"
    }
    $ItemId = (Get-MgDriveItem -DriveId $DriveId -DriveItemId $Path -ErrorAction SilentlyContinue).Id
    if ($ItemId -ne $null) {
        Write-Host "Directory already exists: $Path" -ForegroundColor Red
        return
    }
    $ParentPath = Split-Path -Path $Path -Parent -Resolve:$false
    $ParentId = "root"
    if ($ParentPath -ne "root") {
        $ParentFolder = Get-MgDriveItem -DriveId $DriveId -DriveItemId "${ParentPath}:"
        $ParentId = $ParentFolder.Id
    }
    $Name = Split-Path -Path $Path -Leaf -Resolve:$false
    $NewFolderParams = @{ Name = $Name;  Folder = @{ "@odata.type" = "microsoft.graph.folder" }  }    
    $Item = New-MgDriveItemChild -DriveId $DriveId -DriveItemId $ParentId -BodyParameter $NewFolderParams -ErrorAction SilentlyContinue
    if ( $Item -eq $null) {
        Write-Host "Failed to create directory: $Path" -ForegroundColor Red
        return
    }
    write-host "Directory created: $Path" -ForegroundColor Green
}

function Remove-OneDriveItem {
    param (
        [string]$DriveId,
        [string]$Path,
        [switch]$IsFile = $false
    )
    $textTypeU = "Folder"
    if($IsFile) { $textTypeU = "File" }
    $textTypeL = $textTypeU.ToLower()
    if (!$Path.StartsWith("root:")) {
        $Path = "root:/$Path"
    }
    $Item = Get-MgDriveItem -DriveId $DriveId -DriveItemId $Path -ErrorAction SilentlyContinue
    if ($Item -eq $null) {
        Write-Host "$textTypeU doesn't exists $Path" -ForegroundColor Red
        return
    }
    $IsFolder = $Item.Folder.ChildCount -ne $null
    if ($IsFolder -and $IsFile) {
        Write-Host "The specified path is a folder, not a file $Path" -ForegroundColor Red
        return
    }
    if (!$IsFolder -and !$IsFile) {
        Write-Host "The specified path is a file, not a folder $Path" -ForegroundColor Red
        return
    }
    try {
        Remove-MgDriveItem -DriveId $DriveId -DriveItemId "${Path}:"
        Write-Host "Successfully deleted ${textTypeL}: $Path" -ForegroundColor Green
    } catch {
        Write-Error "Failed to delete the $textTypeL. Verify the path is correct. Error: $_"
    }
}

function List-OneDriveFolder {
    param (
        [string]$DriveId,
        [string]$Path
    )
    if (!$Path.StartsWith("root:")) {
        $Path = "root:/$Path"
    }
    $ItemId = (Get-MgDriveItem -DriveId $DriveId -DriveItemId $Path).Id
    $Children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ItemId
    $countFiles = 0
    $countFolders = 0
    foreach ($Item in $Children) {
        # Check if it's a folder using the sub-property validation
        $IsFolder = $Item.Folder.ChildCount -ne $null
        $fname = $Item.Name
        if ($IsFolder) {
            $fname = $Item.Name
            if ( $fname.Length -gt 87 ) {
                $fname = $fname.Substring(0, 84) + "..."
            } else {
                $fname = $fname.PadRight(87)
            }
            $fmtNbr = $Item.Folder.ChildCount.ToString("N0").PadLeft(15)
            Write-Host "$fname $fmtNbr items" -ForegroundColor Cyan
            $countFolders++
        } else {
            $fname = $Item.Name
            if ( $fname.Length -gt 60 ) {
                $fname = $fname.Substring(0, 57) + "..."
            } else {
                $fname = $fname.PadRight(60)
            }
            $fmtNbr = $Item.Size.ToString("N0").PadLeft(15)
            Write-Host "$fname $($Item.LastModifiedDateTime.ToString())`t$fmtNbr bytes" -ForegroundColor Gray
            $countFiles++
        }
    }
    write-host "`n$Path`nTotal: $countFolders folders, $countFiles files" -ForegroundColor Green
}
function Download-OneDriveFileToLocal {
    param (
        [string]$DriveId,
        $Item,
        [string]$LocalFilePath,
        [switch]$Force = $false
    )
    if (!$Force -and (Test-Path -Path $LocalFilePath)) {
        $LastModified = (Get-Item -Path $LocalFilePath).LastWriteTime.ToString("o")
        if ($LastModified -gt $Item.LastModifiedDateTime.ToString("o")) {
            Write-Host "[FILE] $($Item.Name) Local file is newer than OneDrive version. Skipping download." -ForegroundColor Red
            return
        }
    }
    Write-Host "[FILE] Downloading: $LocalFilePath ($($Item.Size) bytes)..." -ForegroundColor Green
    $null = Get-MgDriveItemContent -DriveId $DriveId -DriveItemId $Item.Id -OutFile $LocalFilePath -Verbose:$false -InformationAction SilentlyContinue -ErrorAction SilentlyContinue
}

function Download-OneDriveFolderToLocal {
    param (
        [string]$DriveId,
        [string]$Path= $null,
        [string]$FolderId = $null,
        [string]$LocalPath = $null,
        [switch]$Force = $false
    )
    if ($Path -ne $null -and $Path.Trim() -ne "" -and !$Path.StartsWith("root:/")) {
        $Path = "root:/$Path"
    }
    if ( $LocalPath -eq $null -or $LocalPath.Trim() -eq "") {
        $LocalPath = Join-Path -Path (Get-Location) -ChildPath (Split-Path -Path $Path -Leaf)
    }
    if ( $Path.Trim() -ne "" ) {
        $srcFolderName = Split-Path -Path $Path -Leaf
        if ( $LocalPath.Contains($srcFolderName) -eq $false ) {
            $LocalPath = Join-Path -Path $LocalPath -ChildPath $srcFolderName
        }
    }
    if ( $FolderId -eq $null -or $FolderId.Trim() -eq "") {
        $Folder = (Get-MgDriveItem -DriveId $DriveId -DriveItemId $Path)
        $FolderId = $Folder.Id
    }
    if (-not (Test-Path -Path $LocalPath)) {
        Write-Host "[FOLDER] Creating: $LocalPath" -ForegroundColor Cyan            
        $null = New-Item -ItemType Directory -Path $LocalPath -Force
    }
    $Children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $FolderId
    foreach ($Item in $Children) {
        $IsFolder = $Item.Folder.ChildCount -ne $null
        $LocalItemPath = Join-Path -Path $LocalPath -ChildPath $Item.Name
        if ($IsFolder) {
            Download-OneDriveFolderToLocal -DriveId $DriveId -FolderId $Item.Id -Path $null -LocalPath $LocalItemPath -Force:$Force
        } else {
            Download-OneDriveFileToLocal -DriveId $DriveId -Item $Item -LocalFilePath $LocalItemPath -Force:$Force
        }
    }
}

function Upload-LocalFileToOneDrive {
    param (
        [string]$DriveId,
        [string]$LocalFilePath,
        [string]$OneDrivePath,
        [switch]$Force = $false
    )
    if ( !(Test-Path -Path $LocalFilePath -PathType Leaf)) {
        Write-Host "Local file not found: $LocalFilePath" -ForegroundColor Red
        return
    }
    if (!$OneDrivePath.StartsWith("root:")) {
        $OneDrivePath = "root:/$OneDrivePath"
    }
    $ShouldUpload = $true
    $TargetFileId = "${OneDrivePath}:"
    try {
        $CloudFile = Get-MgDriveItem -DriveId $DriveId -DriveItemId $TargetFileId -ErrorAction SilentlyContinue
        if ( $CloudFile -ne $null) {
            write-host "CloudFile is not null. CloudFile.Id "  $CloudFile.Id ", Name: " $CloudFile.Name 
            $CloudModifiedDate = [DateTime]::Parse($CloudFile.LastModifiedDateTime)
            $LocalModifiedDate = (Get-Item $LocalFilePath).LastWriteTime
            $TargetFileId = $CloudFile.Id
            if (!$Force -and $LocalModifiedDate -le $CloudModifiedDate) {
                Write-Host "[FILE] $($Item.Name) OneDrive file is newer. $LocalModifiedDate <= $CloudModifiedDate" -ForegroundColor DarkGray
                $ShouldUpload = $false
            }
        }
    } catch {
        # 404 Exception means file doesn't exist yet on the cloud, upload required
    }

    if ($ShouldUpload) {
        $OldProgress = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Write-Host "[FILE] Upload $LocalFilePath to $OneDrivePath..." -ForegroundColor DarkYellow                    
            $SessionParams = @{ item = @{ "@microsoft.graph.conflictBehavior" = "replace" } }
            $UploadSession = New-MgDriveItemUploadSession -DriveId $DriveId -DriveItemId $TargetFileId -BodyParameter $SessionParams 
            $FileStream = New-Object System.IO.FileStream($LocalFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read )
            $StreamLength = $Item.Length
            $null = Invoke-WebRequest -Method PUT -Uri $UploadSession.UploadUrl -Body $FileStream -SkipHeaderValidation -Headers @{
                "Content-Length" = $StreamLength
                "Content-Range"  = "bytes 0-$($StreamLength - 1)/$StreamLength"
            }                    
            $FileStream.Close()
        }
        catch {
            Write-Warning "[FILE] $($Item.Name) Upload failed. Error: $_"
            if ($FileStream) { $FileStream.Close() }
        }
        finally {
            $ProgressPreference = $OldProgress
        }
    }
}
function Upload-LocalFolderToOneDrive {
    param (
        [string]$DriveId,
        [string]$LocalPath,
        [string]$Path,
        [switch]$Force
    )
    if ($Path -ne $null -and $Path.Trim() -ne "" -and !$Path.StartsWith("root:/")) {
        $Path = "root:/$Path"
    }
    $LocalItems = Get-ChildItem -Path $LocalPath
    foreach ($Item in $LocalItems) {
        $OneDriveItemPath = "$Path/$($Item.Name)"
        if ($Item.PSIsContainer) {
            $cloudFolder = Get-MgDriveItem -DriveId $DriveId -DriveItemId "${OneDriveItemPath}:" -ErrorAction SilentlyContinue
            if ( $cloudFolder -eq $null) {
                Write-Host "[FOLDER] Creating: $OneDriveItemPath" -ForegroundColor Yellow                
                $ParentPath = Split-Path -Path $OneDriveItemPath -Parent -Resolve:$false
                $ParentId = "root"
                if ($ParentPath -ne "root") {
                    $ParentFolder = Get-MgDriveItem -DriveId $DriveId -DriveItemId "${ParentPath}:"
                    $ParentId = $ParentFolder.Id
                }
                $NewFolderParams = @{ Name = $Item.Name; Folder = @{ "@odata.type" = "microsoft.graph.folder" } }
                $null = New-MgDriveItemChild -DriveId $DriveId -DriveItemId $ParentId -BodyParameter $NewFolderParams -ErrorAction SilentlyContinue
            }
            Upload-LocalFolderToOneDrive -DriveId $DriveId -LocalPath $Item.FullName -Path $OneDriveItemPath -Force:$Force
        } else {
            Upload-LocalFileToOneDrive -DriveId $DriveId -LocalFilePath $Item.FullName -OneDrivePath $OneDriveItemPath -Force:$Force
        }
    }
}

function cmd-login {
    Connect-MgGraph -TenantId "consumers" -Scopes "Files.ReadWrite", "Files.ReadWrite.All" -ContextScope Process -NoWelcome
    $Drive = Get-MgDrive
    $DriveId = ($Drive | where-object {$_.Name -eq "OneDrive" }).Id
    write-host "Login successful. Drive ID: $DriveId" -ForegroundColor Green
    return $DriveId
}
function cmd-logoff {
    Disconnect-MgGraph -Context (Get-MgContext -All)
    $DriveId = ""
}
function cmd-dir {
    param (
        [string[]]$args
    )
    $path = $cwd
    if ($args.Length -ge 1) {
        $path = $args[0]
    }
    if ( !$path.StartsWith("/") -and !$path.StartsWith("root:") ) {
        $path = "$cwd/$path"
    }
    List-OneDriveFolder -DriveId $DriveId -Path $path
}

function cmd-cd {
    param (
        [string[]]$args
    )
    if ($args.Length -lt 1) {
        Write-Host "Usage: cd <path>" -ForegroundColor Red
        return ""
    }
    $newPath = $args[0]
    if ( $newPath -eq "/") {
        return "root:"
    }
    if ($newPath -eq ".") {
        $newPath = $cwd
    } elseif ( $newPath -eq "..") {
        $newPath = (Split-Path -Path $cwd -Parent).Replace("\", "/")
        if ($newPath -eq "") {
            $newPath = "root"
        }
    } else {
        if (!$newPath.StartsWith("/")) {
            $newPath = "$cwd/$newPath"
        }
    }
    # no need to check just root - it always exists
    if ($newPath -ne "root:" ) {
        $path = $newPath
        if (!$path.StartsWith("root:")) {
            $path = "root:/$path"
        }
        try {
            $ItemId = (Get-MgDriveItem -DriveId $DriveId -DriveItemId $path -ErrorAction SilentlyContinue).Id
            if ($ItemId -eq $null) {
                Write-Host "Directory not found: $newPath" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Directory not found: $newPath" -ForegroundColor Red
            return $newPath
        }
    }
    return $newPath
}
function cmd-md {
    param (
        [string[]]$args
    )
    if ($args.Length -lt 1) {
        Write-Host "Usage: md <path>" -ForegroundColor Red
        return ""
    }
    $newPath = $args[0]
    if ($newPath.StartsWith(".") -or $newPath.StartsWith("..")) {
        Write-Host "Relative path not allowed" -ForegroundColor Red
        return ""
    }
    if (!$newPath.StartsWith("/")) {
        $newPath = "$cwd/$newPath"
    }
    $null = Create-OneDriveFolder -DriveId $DriveId -Path $newPath
}

function cmd-delete {
    param (
        [string]$Path,
        [switch]$IsFile = $false,
        [switch]$Force = $false
    )
    if ($Path.StartsWith(".") -or $Path.StartsWith("..")) {
        Write-Host "Relative path not allowed" -ForegroundColor Red
        return ""
    }
    if (!$Path.StartsWith("/")) {
        $Path = "$cwd/$Path"
    }
    $textType = "folder"
    if ( $IsFile) { $textType = "file" }
    $answer = (Read-Host "Confirm removing $textType $Path by typing [Y]es or [N]o").ToLower()
    if ( !("yes","y" -contains $answer) ) {
        return
    }
    $null = Remove-OneDriveItem -DriveId $DriveId -Path $Path -IsFile:$IsFile -Force:$Force

}
function cmd-rd {
    param (
        [string[]]$args
    )
    if ($args.Length -lt 1) {
        Write-Host "Usage: rd <path>" -ForegroundColor Red
        return ""
    }
    $Force = $args -contains "-f" -or $args -contains "--force"
    cmd-delete -Path $args[0] -IsFile:$false -Force:$Force
}

function cmd-del {
    param (
        [string[]]$args
    )
    if ($args.Length -lt 1) {
        Write-Host "Usage: rd <path>" -ForegroundColor Red
        return ""
    }
    write-host $args[0]
    write-host $args[1]
    $Force = $args -contains "-f" -or $args -contains "--force"
    cmd-delete -Path $args[0] -IsFile:$true -Force:$Force
}

function cmd-download {
    param (
        [string[]]$args
    )
    if ($args.Length -lt 1) {
        Write-Host "Usage: download <source path> (<destination path>)" -ForegroundColor Red
        return ""
    }
    $Path = $args[0]
    $Name = Split-Path -Path $Path -Leaf -Resolve:$false
    $LocalPath = "./$Name"
    if ($args.Length -ge 2 -and !$args[1].StartsWith("-")) {
        $LocalPath = $args[1]
    }
    if ($Path.StartsWith(".") -or $Path.StartsWith("..")) {
        Write-Host "Relative path not allowed" -ForegroundColor Red
        return ""
    }
    if (!$Path.StartsWith("/")) {
        $Path = "$cwd/$Path"
    }
    if ( $LocalPath -eq ".") {
        $LocalPath = (Get-Location).Path
    }
    if ( $LocalPath.StartsWith("./") ) {
        $LocalPath = (Get-Location).Path + "\" + $LocalPath.Substring(2)
    }
    $odPath = $Path
    if (!$odPath.StartsWith("root:")) {
        $odPath = "root:/$odPath"
    }
    $Item = Get-MgDriveItem -DriveId $DriveId -DriveItemId $Path -ErrorAction SilentlyContinue
    if ( $Item -eq $null) {
        Write-Host "Source path not found: $Path" -ForegroundColor Red
        return
    }
    $Force = $args -contains "-f" -or $args -contains "--force"
    #write-host "Path: $Path, LocalPath: $LocalPath, IsFolder: $($Item.Folder.ChildCount -ne $null)"
    if ( $Item.Folder.ChildCount -eq $null) {
        Download-OneDriveFileToLocal -DriveId $DriveId -Item $Item -LocalFilePath $LocalPath -Force:$Force
    } else {
        Download-OneDriveFolderToLocal -DriveId $DriveId -Path $Path -LocalPath $LocalPath -Force:$Force
    }
}

function cmd-upload {
    param (
        [string[]]$args
    )
    if ($args.Length -lt 1) {
        Write-Host "Usage: upload <source path> (<destination path>)" -ForegroundColor Red
        return ""
    }
    $LocalPath = $args[0]
    $IsFolder = $False
    if (Test-Path -Path $LocalPath -PathType Leaf) {
        $IsFolder = $false
    } elseif (Test-Path -Path $LocalPath -PathType Container) {
        $IsFolder = $true
    } else {
        Write-host "The path does not exist: $LocalPath" -ForegroundColor Red
        return
    }    
    $Item = Get-Item -Path $LocalPath
    #$Name = Split-Path -Path $LocalPath -Leaf -Resolve:$false
    $LocalPath = $Item.FullName
    $Path = "$cwd/$($Item.Name)"
    if ($args.Length -ge 2 -and !$args[1].StartsWith("-")) {
        $Path = $args[1]
    }
    if ($Path.StartsWith(".") -or $Path.StartsWith("..")) {
        Write-Host "Relative path not allowed" -ForegroundColor Red
        return ""
    }
    if (!$Path.StartsWith("/") -and !$Path.StartsWith("root:")) {
        $Path = "$cwd/$Path"
    }
    $odPath = $Path
    if (!$odPath.StartsWith("root:")) {
        $odPath = "root:/$odPath"
    }
    $Force = $args -contains "-f" -or $args -contains "--force"
    if ( $IsFolder -eq $false) {
        Upload-LocalFileToOneDrive -DriveId $DriveId -LocalFilePath $LocalPath -OneDrivePath $odPath -Force:$Force
    } else {
        Upload-LocalFolderToOneDrive -DriveId $DriveId -LocalPath $LocalPath -Path $odPath -Force:$Force
    }
}

function odsh {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$False, Position=0)][Alias("cmd")][string]$Command = ""
    )
Write-Host "+========================================+" -ForegroundColor Green
Write-Host "|  OneDrive PowerShell Command Prompt    |" -ForegroundColor Green
Write-Host "|  Type 'help' for list of commands.     |" -ForegroundColor Green
Write-Host "|  Type 'exit' to quit the custom prompt.|" -ForegroundColor Green
Write-Host "+========================================+" -ForegroundColor Green
Write-Host ""

$promptForInput = $true
if ( $Command.Trim() ) {
    $promptForInput = $false
    $InputLine = $Command
}
# Run the command loop indefinitely
while ($true) {    
    if ( $promptForInput ) {
        $InputLine = Read-Host -Prompt "OD:$cwd>"
    }
    $Command = $InputLine.Trim()
    $promptForInput = $true

    if ($Command -eq "exit") {
        Write-Host "Exiting custom command prompt. Goodbye!" -ForegroundColor Yellow
        break
    }

    if ([string]::IsNullOrWhiteSpace($Command)) {
        continue
    }

    $args = $Command.Split(" ", [StringSplitOptions]::RemoveEmptyEntries)
    if ( $Command.Contains(" ") ) {
        $Pattern = '"([^"]*)"|''([^'']*)''|([^\s]+)'
        $args = [regex]::Matches($Command, $Pattern) | ForEach-Object {
            if ($_.Groups[1].Value) { $_.Groups[1].Value }
            elseif ($_.Groups[2].Value) { $_.Groups[2].Value }
            else { $_.Groups[0].Value }
        }
    }

    <#
    Write-Host "Total elements parsed: $($args.Count)`n" -ForegroundColor Green
    for ($i = 0; $i -lt $args.Count; $i++) {
        Write-Host "args[$i]: $($args[$i])" -ForegroundColor Cyan
    }
    #>
    #continue
    switch -Regex ($args[0].ToLower()) {
        "^(help|\?|--help|-h)$" { print-help -Args $args }
        "^(login)$"             { $DriveId = cmd-login  }
        "^(logoff)$"            { cmd-logoff }
        "^(cd)$"                { if ($args.Length -ge 2) {
                                     $newPath = cmd-cd -args @args
                                     if ($newPath -ne "") { 
                                        $cwd = $newPath 
                                     } 
                                  }
                                }
        "^(dir|ls)$"            { cmd-dir -args @args }
        "^(md|mkdir)$"          { if ($args.Length -ge 2) { cmd-md -args @args } }
        "^(rd|rmdir)$"          { if ($args.Length -ge 2) { cmd-rd -args @args } }
        "^(del|rm)$"            { if ($args.Length -ge 2) { cmd-del -args @args } }
        "^(download|dl)$"       { if ($args.Length -ge 2) { cmd-download -args @args } }
        "^(upload|ul)$"         { if ($args.Length -ge 2) { cmd-upload -args @args } }
    }
    if ( $Command.StartsWith("!")) {
         $Command = $Command.Substring(1)
        try {
            # Invoke-Expression runs the text string as a functional PowerShell command
            Invoke-Expression -Command $Command
        }
        catch {
            # Catch syntax errors or missing cmdlet issues cleanly without crashing the loop
            Write-Error "Error executing command: $_"
        }
    }
}

}

Export-ModuleMember -Function odsh