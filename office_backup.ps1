


$macroPath = "$env:APPDATA\Microsoft\Excel\XLSTART"
$savePath  = 'C:\Development_Solutions\backup\office-configs\xl'
$buItems   = gci $macroPath

foreach ($item in $buItems){  
    $item | Copy-Item -Destination "$savePath\backup.PERSONAL.XLSB" -Verbose
}










#$savePath = "c:"+$env:HOMEPATH+'\OneDrive - Serco\Documents\backup\office-configs\xl'

