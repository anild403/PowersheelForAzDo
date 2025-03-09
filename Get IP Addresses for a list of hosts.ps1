$hostnames = Get-Content "C:\temp\hostnames.txt"
$results = foreach ($hostname in $hostnames) {
    $ping = Test-Connection -ComputerName $hostname -Count 1 -ErrorAction SilentlyContinue
    if ($ping) {
        [PSCustomObject]@{
            Hostname = $hostname
            IPAddress = $ping.IPV4Address.IPAddressToString
        }
    }
}

$results | Export-Csv -Path "C:\temp\output.csv" -NoTypeInformation