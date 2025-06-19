# OneDrive Killer
# This script will stop all OneDrive processes, remove the global OneDrive installation

[console]::OutputEncoding = [System.Text.Encoding]::UTF8
[console]::InputEncoding = [System.Text.Encoding]::UTF8

# 1. Check if the script is running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Administrator)) {
    Write-Host "管理者権限で起動してください！" -ForegroundColor Red
    exit 0
}

Write-Host "管理者権限を動き中..." -ForegroundColor Green

function Get-OneDrivePath {
    $onedrivePath = $env:OneDrive
    if ($onedrivePath -eq $null) {
        $onedrivePath = $env:OneDriveConsumer
    }
    return $onedrivePath
}

# 2. Check if OneDrive has taken over the user's folder
function Test-OneDriveTakeover {
    $documentsPath = [Environment]::GetFolderPath("MyDocuments")
    $picturesPath = [Environment]::GetFolderPath("MyPictures")
    $musicPath = [Environment]::GetFolderPath("MyMusic")
    $videosPath = [Environment]::GetFolderPath("MyVideos")
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    
    $onedrivePath = Get-OneDrivePath
    
    $isTakeover = $false
    
    # Check if the documents, music, pictures, videos, and desktop are pointing to OneDrive
    foreach ($path in @($documentsPath, $musicPath, $picturesPath, $videosPath, $desktopPath)) {
        if ($path -like "*OneDrive*") {
            $isTakeover = $true
            Write-Host "OneDriveはユーザーのフォルダーを管理しています: $path" -ForegroundColor Yellow
        }
    }
    
    return $isTakeover
}

# Check if OneDrive is taking over the user's folder
if (-not (Test-OneDriveTakeover)) {
    Write-Host "OneDriveはユーザーのフォルダーを管理していません。スクリプトを終了します..." -ForegroundColor Green
    exit 0
}

# 3. Stop OneDrive processes and restore folders
function Stop-OneDriveProcesses {
    Write-Host "OneDriveプロセスを終了しています..." -ForegroundColor Yellow
    
    $onedriveProcesses = @(
        "OneDrive",
        "OneDriveSetup",
        "OneDriveStandaloneUpdater"
    )
    
    foreach ($processName in $onedriveProcesses) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($processes) {
            Write-Host "プロセスを終了しています: $processName" -ForegroundColor Yellow
            Stop-Process -Name $processName -Force -ErrorAction SilentlyContinue
        }
    }
    
    # Wait for the processes to fully end
    Start-Sleep -Seconds 2
}

function Check-LockedItem {
    param([string]$path)
    $name = Split-Path "$path" -leaf
    $pathorg = Split-Path "$path" -Parent
    try {
        Rename-Item -Path "$path" -NewName "$name--" -ErrorAction Stop
    }
    catch {
        return $true
    }
    finally {
        Rename-Item -path "$pathorg\$name--" -NewName $name -ErrorAction SilentlyContinue
    }
    return $false
}

function Set-KnownFolderPath {
    Param (
            [Parameter(Mandatory = $true)]
            [ValidateSet('Documents', 'Pictures', 'Music', 'Videos', 'Desktop')]
            [string]$KnownFolder,
            [Parameter(Mandatory = $true)]
            [string]$Path
    )
    
    # Define known folder GUIDs
    $KnownFolders = @{
        'Desktop' = 'B4BFCC3A-DB2C-424C-B029-7FE99A87C641';
        'Documents' = 'FDD39AD0-238F-46AF-ADB4-6C85480369C7';
        'Pictures' = '33E28130-4E1E-4676-835A-98395C3BC3BB';
        'Music' = '4BD8D571-6D19-48D3-BE97-422220080E43';
        'Videos' = '18989B1D-99B5-455B-841C-AB7C74E4DDFC';
    }
    
    $Type = ([System.Management.Automation.PSTypeName]'KnownFolders').Type
    if (-not $Type) {
        $Signature = @'
[DllImport("shell32.dll")]
public extern static int SHSetKnownFolderPath(ref Guid folderId, uint flags, IntPtr token, [MarshalAs(UnmanagedType.LPWStr)] string path);
'@
        $Type = Add-Type -MemberDefinition $Signature -Name 'KnownFolders' -Namespace 'SHSetKnownFolderPath' -PassThru
    }
    
    # Validate the path
    if (Test-Path $Path -PathType Container) {
        # Call SHSetKnownFolderPath
        $ret = $Type::SHSetKnownFolderPath([ref]$KnownFolders[$KnownFolder], 0, 0, $Path)
        if($ret -ne 0) {
            throw New-Object System.Exception "Failed to set known folder path. Error code: $ret"
        }
    } else {
        throw New-Object System.IO.DirectoryNotFoundException "Could not find part of the path $Path."
    }
}

function Restore-OneDriveFolders {
    $userProfile = $env:USERPROFILE
    $onedrivePath = Get-OneDrivePath
    $folders = @("Documents", "Pictures", "Music", "Videos", "Desktop")

    foreach ($folder in $folders) {
        $originalPath = "$userProfile\$folder"
        $currentPath = "$onedrivePath\$folder"
        if (!(Test-Path $currentPath)) {
            continue
        }
        if(Test-Path $originalPath) {
            Write-Host "$originalPath フォルダーが存在します。ファイルを移動しています..." -ForegroundColor Green
            Get-ChildItem -Path $currentPath -Recurse | Move-Item -Destination $originalPath
            Set-KnownFolderPath -KnownFolder $folder -Path $originalPath
        } else {
            if(Check-LockedItem $currentPath) {
                Write-Host "$currentPath フォルダーがロックされています。パソコンを再起動してください。" -ForegroundColor Red
                exit 0
            } else {
                Write-Host "$currentPath フォルダーを移動しています..." -ForegroundColor Green
                Move-Item -Path $currentPath -Destination $originalPath
            }
        }
    }
}

Stop-OneDriveProcesses
Restore-OneDriveFolders

# 4. uninstall OneDrive
function Uninstall-OneDrive {
    Write-Host "OneDriveをアンインストールしています..." -ForegroundColor Yellow
    
    # find OneDrive installation path
    $onedrivePaths = @(
        "C:\Windows\System32\OneDriveSetup.exe",
        "${env:ProgramFiles}\Microsoft OneDrive\OneDriveSetup.exe",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDriveSetup.exe"
    )
    
    foreach ($setupPath in $onedrivePaths) {
        if (Test-Path $setupPath) {
            try {
                # use OneDriveSetup.exe to uninstall
                Start-Process -FilePath $setupPath -ArgumentList "/uninstall /allusers" -Wait -NoNewWindow
                Write-Host "OneDriveをアンインストールしました。" -ForegroundColor Green
                break
            }
            catch {
                Write-Host "OneDriveをアンインストールできませんでした: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    
    # use Windows feature to uninstall
    try {
        Write-Host "OneDriveをWindows機能を使用してアンインストールしています..." -ForegroundColor Yellow
        $null = Get-WindowsOptionalFeature -Online -FeatureName "OneDriveSyncClient" | Disable-WindowsOptionalFeature -Online -NoRestart
        Write-Host "OneDrive Windows機能を無効にしました。" -ForegroundColor Green
    }
    catch {
        Write-Host "OneDriveをWindows機能を使用してアンインストールできませんでした。" -ForegroundColor Red
    }
}

Uninstall-OneDrive

function Set-OneDriveRegistrySettings {
    Write-Host "レジストリ項目を設定してOneDriveが再接続しないようにしています..." -ForegroundColor Yellow
    
    # disable file sync
    $path = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive"
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    Set-ItemProperty -Path $path -Name "DisableFileSyncNGSC" -Value 1 -Type DWord -Force
    Write-Host "レジストリ項目: $path\DisableFileSyncNGSC = 1" -ForegroundColor Green
    
    # hide OneDrive
    $namespacePath = "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    try {
        if (-not (Test-Path $namespacePath)) {
            New-Item -Path $namespacePath -Force | Out-Null
        }
        Set-ItemProperty -Path $namespacePath -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Type DWord -Force
        Write-Host "OneDriveを非表示にしました: $namespacePath\System.IsPinnedToNameSpaceTree = 0" -ForegroundColor Green
    }
    catch {
        Write-Host "OneDriveを非表示にできませんでした" -ForegroundColor Red
    }
    
    # disable OneDrive auto start
    $runPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    try {
        Remove-ItemProperty -Path $runPath -Name "OneDrive" -Force -ErrorAction SilentlyContinue
        Write-Host "OneDrive自動起動項目を削除しました。" -ForegroundColor Green
    }
    catch {
        Write-Host "OneDrive自動起動項目を削除できませんでした。" -ForegroundColor Red
    }
}

Set-OneDriveRegistrySettings

function Restart-Explorer {
    Write-Host "リソースマネージャーを再起動して変更を適用しています..." -ForegroundColor Yellow
    try {
        Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process "explorer"
        Write-Host "リソースマネージャーを再起動しました。" -ForegroundColor Green
    }
    catch {
        Write-Host "リソースマネージャーを再起動できませんでした。" -ForegroundColor Red
    }
}

Restart-Explorer