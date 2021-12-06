$getFirstLine = $true

get-childItem ".\*.txt" | ForEach-Object {
    $filePath = $_

    $lines =  $lines = Get-Content $filePath  
    $linesToWrite = switch($getFirstLine) {
           $true  {$lines}
           $false {$lines | Select-Object -Skip 1}

    }

    $getFirstLine = $false
    Add-Content ".\5.txt" $linesToWrite
    }