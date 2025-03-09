$hostnames = Get-Content "C:\temp\hostnames.txt"
$results = foreach ($hostname in $hostnames) {
$items = $hostname.Split(",")
$computer = $items[0]
$port = $items[1]
    $result = Test-NetConnection -ComputerName $computer -Port $port
     if ($result.TcpTestSucceeded) {
        Write-Host $computer ": Telnet works"
    } else {
        Write-Host $computer ": Telnet failed"
    }
}
