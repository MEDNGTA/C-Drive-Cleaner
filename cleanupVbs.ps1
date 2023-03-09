param([switch]$Elevated)

function Prev-Admin {
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Prev-Admin) -eq $false) {
if($elevated)
{
#--------------------------
}
else {
Start-Process powershell.exe -Verb runAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($MyInvocation.MyCommand.Definition))
}

exit
}
cscript $PSScriptRoot"\cleanup.vbs"
#Clean chrome cache
