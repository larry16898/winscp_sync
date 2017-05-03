# Load WinSCP .NET assembly
#Add-Type -Path "WinSCPnet.dll"
 
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
param(
    [string]$HostName='192.168.56.102',
    [string]$UserName='updateuser',
    [string]$Password='abc@123',
    [string]$SshHostKeyFingerprint='ssh-rsa 2048 d9:a9:62:39:33:41:30:06:13:cb:3b:86:e8:e3:da:33',
    [string]$local_path='C:\Program Files\tbtools',
    [string]$remote_path='/updatesource/cs_office'
)
    try
    {
        # Load WinSCP .NET assembly
        Add-Type -Path "WinSCPnet.dll"

        # determine if the local destination folder is exists.
        $local_path = 'C:\Program Files\tbtools'
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

function CreateShortCut
{
    param(
        [string]$shortcutName,
        [string]$programName,
        [string]$installPath
    )
    try
    {
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("$home\desktop\$shortcutName")
    $Shortcut.TargetPath = Join-Path $installPath $programName
    $Shortcut.WorkingDirectory = $installPath
    $Shortcut.Save()
    }
    catch [Exception]
    {
        Write-Host ("Error: {0}" -f $_.Exception.Message)
        exit 1
    }
}

# write my own code here
$tb_tools_list = @(@{programName='skype_helper.exe';installPath='C:\Program Files\tbtools\skype_helper';shortcutName='客服Skype登录助手.lnk'})

SyncStuff -HostName '192.168.56.102' -UserName 'updateuser' -Password 'abc@123' -SshHostKeyFingerprint 'ssh-rsa 2048 d9:a9:62:39:33:41:30:06:13:cb:3b:86:e8:e3:da:33' `
                -local_path 'C:\Program Files\tbtools' -remote_path '/updatesource/cs_office'

foreach ($tool in $tb_tools_list)
{
    CreateShortCut -shortcutName $tool.shortcutName -programName $tool.programName -installPath $tool.installPath
}