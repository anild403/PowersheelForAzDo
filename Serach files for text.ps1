
$AllFiles = Get-ChildItem –recurse
$VBPFiles = $AllFiles | where {$_.extension –in “.sql”, ".bas", ".frm" }
$a = $VBPFiles | Select-String "GET_UNAPPROVED_PAYMENTS"# "Update " Pattern 'gt*.ocx' -List | Select Path



 $a | foreach {
 $path = $_.Path 
 Write-Host  $_.Path 
 get-content $_.Path -ReadCount 10000 |
 foreach { $_ -match "GET_UNAPPROVED_PAYMENTS" }  
     if($_ -match "GET_UNAPPROVED_PAYMENTS"){
     $text = $_.Path+"~"+$_.Line  
     Add-Content C:\temp\Oracle.csv $text  
       
    }
}