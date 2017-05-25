#[CmdletBinding()]
#param(
$HostName='172.19.12.6'
$UserName='updateuser'
$Password='U1ikKnMvmE'
$SshHostKeyFingerprint='ssh-rsa 2048 e5:90:74:6b:dd:a8:cc:f4:15:a9:41:3c:d4:48:4e:53'
$ParentFolderName='office_tools'
$SubFolderName='skype_helper_cs'
$AddShortcutName="客服skype登录小助手.lnk"      #桌面快捷方式文件名，后缀名.lnk, 如果字段不为空，则创建快捷方式
$ProgramName="skype_helper.exe"        #可执行程序文件名，后缀名.exe
$RemoveShortcuts="skype登陆器.lnk","Skype登录器.lnk","客服skype登录助手.lnk"    #需要删除的桌面快捷方式文件名， 字符串数组，命令行参数以逗号分隔
$IsSyncFirst = "1"           #   1：  先同步再启动     其它值：   直接从本地启动
$FileMask=""                                          #  同步文件掩码，  例如 |Log\
#)

$LOCAL_ROOT_PATH = 'C:\Program Files\tbtools'
$LOCAL_INSTALLED_PATH = Join-Path (Join-Path $LOCAL_ROOT_PATH $ParentFolderName) $SubFolderName
$REMOTE_ROOT_PATH = '/updatesource/client_exe'
 
# Session.FileTransferred event handler
function FileTransferred
{
    param($e)
 
    if ($e.Error -eq $Null)
    {
        Write-Host ("Update of {0} succeeded" -f $e.FileName) -ForegroundColor green
    }
    else
    {
        Write-Host ("Update of {0} failed: {1}" -f $e.FileName, $e.Error) -ForegroundColor Red
    }
 
    if ($e.Chmod -ne $Null)
    {
        if ($e.Chmod.Error -eq $Null)
        {
            Write-Host ("Permisions of {0} set to {1}" -f $e.Chmod.FileName, $e.Chmod.FilePermissions)
        }
        else
        {
            Write-Host ("Setting permissions of {0} failed: {1}" -f $e.Chmod.FileName, $e.Chmod.Error)
        }
 
    }
    else
    {
        Write-Host ("Permissions of {0} kept with their defaults" -f $e.Destination)
    }
 
    if ($e.Touch -ne $Null)
    {
        if ($e.Touch.Error -eq $Null)
        {
            Write-Host ("Timestamp of {0} set to {1}" -f $e.Touch.FileName, $e.Touch.LastWriteTime)
        }
        else
        {
            Write-Host ("Setting timestamp of {0} failed: {1}" -f $e.Touch.FileName, $e.Touch.Error)
        }
 
    }
    else
    {
        # This should never happen during "local to remote" synchronization
        Write-Host ("Timestamp of {0} kept with its default (current time)" -f $e.Destination)
    }
}

# Session.FileTransferProgress event handler
function FileTransferProgress
{
    param($e)
 
    # New line for every new file
    if (($script:lastFileName -ne $Null) -and
        ($script:lastFileName -ne $e.FileName))
    {
        Write-Host
    }
 
    # Print transfer progress
    Write-Host -NoNewline ("`r{0} ({1:P0})" -f $e.FileName, $e.FileProgress)
 
    # Remember a name of the last file reported
    $script:lastFileName = $e.FileName
}

function SyncStuff
{
    try
    {
        # Load WinSCP .NET assembly
        Add-Type -Path "WinSCPnet.dll"

        # determine if the local destination folder is exists.
        $local_path = $LOCAL_INSTALLED_PATH
        $remote_path = $REMOTE_ROOT_PATH + "/$ParentFolderName/$SubFolderName"
        if(!(Test-Path $local_path))
        {
            New-Item -Path $local_path -ItemType Directory -ErrorAction Stop >$null
        }
 
        # Setup session options
        $sessionOptions = New-Object WinSCP.SessionOptions -Property @{
            Protocol = [WinSCP.Protocol]::Sftp
            HostName = $HostName
            UserName = $UserName
            Password = $Password
            SshHostKeyFingerprint = $SshHostKeyFingerprint
        }

        # Setup TransferOptions
        $transferOptions = New-Object WinSCP.TransferOptions
        $transferOptions.FileMask = $FileMask
 
        $session = New-Object WinSCP.Session

        try
        {
            # Will continuously report progress of synchronization
            #$session.add_FileTransferred( { FileTransferred($_) } )

            $session.add_FileTransferProgress( { FileTransferProgress($_) } )
 
            # Connect
            $session.Open($sessionOptions)

            Write-Host -ForegroundColor Green "亲，俺正在准备中，第一次同步可能会比较慢，请耐心等待..."
 
            # Synchronize files using Mirror mode
            $synchronizationResult = $session.SynchronizeDirectories(
                [WinSCP.SynchronizationMode]::Local, $local_path, $remote_path, $True, $True,[WinSCP.SynchronizationCriteria]::Time, $transferOptions)
 
            # Throw on any error
            $synchronizationResult.Check()

            Write-Host -ForegroundColor Green -BackgroundColor Black "同步成功完成！"
        }
        finally
        {
            # Disconnect, clean up
            $session.Dispose()
        }

 #       exit 0
    }
    catch [Exception]
    {
        Write-Host ("Error: {0}" -f $_.Exception.Message)
        Write-Host "请按任意键关闭此窗口！"
        $void = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }

}

function CreateDesktopShortCut
{
    $installPath = $LOCAL_INSTALLED_PATH
#    Write-Host -ForegroundColor Red $installPath
    try
    {
        $WshShell = New-Object -ComObject WScript.Shell
#        Write-Host -ForegroundColor Red "WshShell is done."
        if(!(Test-Path "$home\desktop\$AddShortcutName"))
        {
#            Write-Host -ForegroundColor Red "not exists $home\desktop\$shortcutName"
            $Shortcut = $WshShell.CreateShortcut("$home\desktop\$AddShortcutName")
            $Shortcut.TargetPath = Join-Path $installPath $ProgramName
            $Shortcut.WorkingDirectory = $installPath
            $Shortcut.Save()
            Write-Host -ForegroundColor Green "创建桌面快捷方式 $AddShortcutName 完成！"
        }
    }
    catch [Exception]
    {
        Write-Host ("Error: {0}" -f $_.Exception.Message)
        exit 1
    }
}

function RemoveOldDesktopShortCut
{
    if($RemoveShortcuts.Length -gt 0) 
    {
        foreach ($shortcut in $RemoveShortcuts)
        {
            $shortcut_filepath = Join-Path "$home\desktop" $shortcut
            if(Test-Path $shortcut_filepath)
            {
                Remove-Item $shortcut_filepath -Force
                Write-Host -ForegroundColor Green "已清理旧的桌面快捷方式 $shortcut"
            }
        }
    }
}
# write my own code here
# start to sync
if ($IsSyncFirst -eq "1")
{
    SyncStuff
}
# remove old shortcut names
RemoveOldDesktopShortCut

# Create shortcut on desktop
if ($AddShortcutName)
{
    CreateDesktopShortCut
}

$local_installed_program_file = Join-Path $LOCAL_INSTALLED_PATH $ProgramName

# stop process if exists.
$ProgramName = Split-Path $ProgramName -Leaf     # e.g. if programeName is "Debug\test.exe", this will return only "test.exe"
Get-Process -Name "$ProgramName" 2>&1 >$null | Stop-Process -Force 2>&1 >$null


Start-Process -FilePath "$local_installed_program_file" -WorkingDirectory "$LOCAL_INSTALLED_PATH"


