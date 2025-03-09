#Install-Module -Name ImportExcel -Force -AllowClobber

# Define the path to the input text file and Excel output file
$inputFilePath = "C:\Projects\My Project\SH\DevAndTesting\CurrentSprintStories.txt"
$outputFilePath = "C:\Projects\My Project\SH\DevAndTesting\BugList.xlsx"

# Read the list of stories from the input text file
$stories = Get-Content -Path $inputFilePath

# Create an empty array to store the bug details
$bugs = @()
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)"))
$headers = @{Authorization=("Basic {0}" -f $base64AuthInfo)}

# Loop through each story and call the Azure DevOps API to retrieve associated bugs
foreach ($story in $stories) {
    # Construct the API URL to retrieve bugs associated with the story
    $apiUrl = "https://yourcompanyurl/_apis/wit/workItems?ids=$story&`$expand=relations&api-version=6.0"

    # Invoke the Azure DevOps API using personal access token
    $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get

    foreach($relation in $response.value.relations){
     $bugId = $relation.url.Split("/")[-1]
        #$bugUrl = "https://yourcompanyurl/_apis/wit/workItems/"+$bugId+"&`$expand=all?api-version=6.0"
        $bugUrl = "https://yourcompanyurl/_apis/wit/workItems?ids=$bugId&`$expand=relations&api-version=6.0"
        $bug = Invoke-RestMethod -Uri $bugUrl -Headers $headers -Method Get
        if($bug.value.fields.'System.WorkItemType' -eq "Bug" -and $story -eq $bug.value.fields.'System.Parent'){
             $bugs += [PSCustomObject]@{
                ParentID = $bug.value.fields.'System.Parent'
                ID = $bug.value.id
                Title = $bug.value.fields.'System.Title'
                ResolvedBy  = $bug.value.fields.'Microsoft.VSTS.Common.ResolvedBy'.displayName
            }
    
        }
    }

   
}

# Export the bug details to an Excel file
$bugs | Export-Excel -Path $outputFilePath -AutoSize -NoHeader -FreezeTopRow

Write-Host "Bug details have been written to '$outputFilePath'."
