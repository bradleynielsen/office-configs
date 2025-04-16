$macroPath = "$env:APPDATA\Microsoft\Excel\XLSTART"
$savePath  = 'C:\Development_Solutions\github\office-configs\xl'
$buItems   = gci $macroPath

foreach ($item in $buItems){  
    $item | Copy-Item -Destination "$savePath\backup.PERSONAL.XLSB" -Verbose
}

