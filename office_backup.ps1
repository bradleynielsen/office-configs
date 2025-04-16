$gitPath  = "C:\Program Files\Git\git-bash.exe"

$rootpath   = $PSScriptRoot
$syncScript = $rootpath+"\sync.sh"



$macroPath = "$env:APPDATA\Microsoft\Excel\XLSTART"
$savePath  = 'C:\Development_Solutions\backup\office-configs\xl'
$buItems   = gci $macroPath

foreach ($item in $buItems){  
    $item | Copy-Item -Destination "$savePath\backup.PERSONAL.XLSB" -Verbose
}




git status
git add .
git commit -m "sync script commit"
git push



<#

#$savePath = "c:"+$env:HOMEPATH+'\OneDrive - Serco\Documents\backup\office-configs\xl'

cd $rootpath

& $gitPath $syncScript

#>