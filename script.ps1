# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Unrestricted
#cargo clean
$var = $args[0]

if($var -eq "s") {
    & 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe' controlvm "Win10" poweroff
    Exit
}

cargo build
Remove-Item .\vm\TextService.dll
Remove-Item .\vm\*.log
Copy-Item .\target\debug\TextService.dll .\vm

& 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe' controlvm "Win10" poweroff
& 'C:\Program Files\Oracle\VirtualBox\VBoxManage.exe' startvm "Win10"
