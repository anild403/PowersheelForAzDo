# Set variables
$organization = "your-organization"
$project = "your-project"
$repository = "workers-comp"
$apiVersion = "6.1-preview.1"
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"
$responseFile = "C:\Users\userid\Downloads\apiresp.txt"


# Calculate date range (last year)
$startDate = Get-Date "04/01/2022"
$endDate = (Get-Date).ToString("yyyy-MM-dd")

# Set API endpoint
$uri = "https://yourcompanyurl/yourproject/_apis/git/repositories/$repository/commits?searchCriteria.itemVersion.version=master&searchCriteria.itemVersion.versionType=branch&searchCriteria.fromDate=$startDate&searchCriteria.toDate=$endDate&`$skip=0&api-version=$apiVersion&`$top=500000"

# Create header with personal access token
$header = @{
    Authorization = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat")))"
}

# Call API and retrieve data
$response = Invoke-RestMethod -Uri $uri -Headers $header -Method Get

# Convert response object to JSON string
$jsonString = ConvertTo-Json $response

# Save JSON string to text file
$jsonString | Out-File -FilePath $responseFile -Encoding utf8
$linesChanged = 0
$linesDeleted = 0
$linesAdded = 0

foreach ($commit in $response.value) {
    $linesChanged += $commit.changeCounts.Edit
    $linesDeleted += $commit.changeCounts.Delete    
    $linesAdded += $commit.changeCounts.Add    
}

$commitCount = $response.count
# Output results
Write-Host "Since $startDate, for  $repository repo, $commitCount commits happened, $linesChanged lines were edited, $linesAdded were added and $linesDeleted lines were deleted in the $repository repository."
