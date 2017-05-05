[CmdletBinding()]
param(
    [Parameter(Mandatory=$True)][string]$HostName='192.168.56.102',
    [Parameter(Mandatory=$True)][string]$UserName='updateuser',
    [Parameter(Mandatory=$True)][string]$Password='abc@123',
    [Parameter(Mandatory=$True)][string]$SshHostKeyFingerprint='ssh-rsa 2048 d9:a9:62:39:33:41:30:06:13:cb:3b:86:e8:e3:da:33',
    [Parameter(Mandatory=$True)][string]$ParentFolderName,
    [Parameter(Mandatory=$True)][string]$SubFolderName,
    [Parameter(Mandatory=$True)][string]$AddShortcutName,      #桌面快捷方式文件名，后缀名.lnk
    [Parameter(Mandatory=$True)][string]$ProgramName,        #可执行程序文件名，后缀名.exe
    [Parameter(Mandatory=$True)][string[]]$RemoveShortcuts    #需要删除的桌面快捷方式文件名， 字符串数组，命令行参数以逗号分隔
)

$LOCAL_ROOT_PATH = 'C:\Program Files\tbtools'
$LOCAL_INSTALLED_PATH = Join-Path (Join-Path $LOCAL_ROOT_PATH $ParentFolderName) $SubFolderName
$REMOTE_ROOT_PATH = '/updatesource'
 
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
 
        $session = New-Object WinSCP.Session

        try
        {
            # Will continuously report progress of synchronization
            $session.add_FileTransferred( { FileTransferred($_) } )
 
            # Connect
            $session.Open($sessionOptions)
 
            # Synchronize files
            $synchronizationResult = $session.SynchronizeDirectories(
                [WinSCP.SynchronizationMode]::Local, $local_path, $remote_path, $True)
 
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
SyncStuff

# remove old shortcut names
RemoveOldDesktopShortCut

# Create shortcut on desktop
CreateDesktopShortCut

Write-Host "请按任意键关闭此窗口！"
$void = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
