
$AzureDevOpsPAT = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"

$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($AzureDevOpsPAT)")) }

$uriAccount = "https://yourcompanyurl/yourproject/_apis/git/repositories/"
$response= Invoke-RestMethod -Uri $uriAccount -Method get -Headers $AzureDevOpsAuthenicationHeader 

$response | foreach {
$_.Value.GetEnumerator() | foreach {
$dictionary = @{}
        $uri = "https://yourcompanyurl/yourproject/_apis/git/repositories/{" + $_.id + "}/commits?`$top=100000&searchCriteria.fromDate=2022-04-01&api-version=6.0"
        $commits= Invoke-RestMethod -Uri $uri -Method get -Headers $AzureDevOpsAuthenicationHeader 
        Write-Host "name: " $_.name ", commit count:"$commits.count

        $commits | foreach {
$_.Value.GetEnumerator() | foreach {
$key = $_.author.name
    if ($dictionary.ContainsKey($key)) {
        $dictionary[$key]++
    } else {
        $dictionary[$key] = 1
    }


#Write-Host $_.author
}}
  foreach ($key in $dictionary.Keys) 
        {
    Write-Host "repo: " $_.name", name: "$key", count :" $($dictionary[$key])
}
        }
        }


      