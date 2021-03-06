Param
     (
         [Parameter(Mandatory=$true)]
         [String]$EXE_Name,    
         [Parameter(Mandatory=$true)]
         [String]$Sources_Folder,            
         [Parameter(Mandatory=$true)]
         [String]$PS1_Name,        
         [Parameter(Mandatory=$true)]
         [String]$Conf_File,            
         [Parameter(Mandatory=$true)]
         [AllowEmptyString()]        
         [String]$Icon_File
     )
     
 $Temp_Folder = $env:TEMP
 $User_Profile = $env:userprofile
 $Global:User_Desktop = "$User_Profile\Desktop"
 $Global:Winrar_Folder = "C:\Program Files\WinRAR"
 $Global:Temp_Conf_file = "$Temp_Folder\My_Conf_File.conf"
 copy-item $Conf_File $Temp_Conf_file
 Add-Content $Temp_Conf_file "Setup=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -sta -noprofile -executionpolicy bypass -file %temp%\$PS1_Name"
 $command = "$Winrar_Folder\WinRAR.exe"    
 & $command a -ep1 -r -o+ -dh -ibck -sfx -iadm "-iicon$Icon_File" "-z$Temp_Conf_file"  "$User_Desktop\$EXE_Name.exe" "$Sources_Folder\*"
 sleep 5
 remove-item $Temp_Conf_file -Force
 write-host "" 
 write-host "******************************************************" -foregroundcolor "cyan"
 write-host "The EXE $EXE_Name has been created on your desktop"     -foregroundcolor "yellow"
 write-host "******************************************************" -foregroundcolor "cyan"

#Following are the different WinRAR switches used in the script:

# a: Add files to archive
 # -ep1: Exclude base directory from names
 # -r: Repair an archive
 # -o+: Overwrite all
 # -dh: Open shared files
 # -ibck: Run Winrar in Background
 # -sfx: Create an SFX self-extracting archive
 # -iadm: request administrative access for SFX archive
 # -iiconC: Specify the icon
 # -z: Read archive comment from file