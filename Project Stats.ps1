$PAT = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"
Param
(
    [string]$PAT 
)

$AzureDevOpsAuthenicationHeader = @{Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($PAT)")) } + @{"Content-Type"="application/json"; "Accept"="application/json"}
$UriOrganization = "https://yourcompanyurl/"
$UriOrganizationRM = "https://yourcompanyurl/"
$featurearr = @()
$filepath = "C:\Users\userid\Desktop\ProjectStats.xlsx"

$monthAgo = (Get-Date).AddMonths(-1).ToString("yyyy-MM-dd")

$uriProject = $UriOrganization + "_apis/projects?`$top=500&api-version=6.1-preview.4"
$ProjectsResult = Invoke-RestMethod -Uri $uriProject -Method get -Headers $AzureDevOpsAuthenicationHeader
Foreach ($project in $ProjectsResult.value)
{
    $uriProjectStats = $UriOrganization + "_apis/Contribution/HierarchyQuery/project/$($project.id)?api-version=6.1-preview.1"   
    $projectStatsBody = @{
        "contributionIds"= @("ms.vss-work-web.work-item-metrics-data-provider-verticals", "ms.vss-code-web.code-metrics-data-provider-verticals", "ms.vss-code-web.build-metrics-data-provider-verticals")
        "dataProviderContext" = @{
            "properties" =@{
                "numOfDays"=30
                "sourcePage"=@{
                    "url"=($UriOrganization + $project.name)
                    "routeId"="ms.vss-tfs-web.project-overview-route"
                    "routeValues" =@{
                        "project" = $project.id
                        "controller"="Apps"
                        "action"="ContributedHub"
                        "serviceHost"=$Organization
                        }
                    }          
                }
            }
        }  | ConvertTo-Json -Depth 5

    $projectStatsResult = Invoke-WebRequest -Uri $uriProjectStats -Headers $AzureDevOpsAuthenicationHeader -Method Post -Body $projectStatsBody 
    $projectStatsJson = ConvertFrom-Json $projectStatsResult.Content

    $workItemsCreated = 0
    $workItemsCompleted = 0
    $commitsPushed = 0
    $pullRequestsCreated = 0
    $pullRequestsCompleted = 0

    $workItemsCreated = $projectStatsJson.dataProviders.'ms.vss-work-web.work-item-metrics-data-provider-verticals'.workMetrics.workItemsCreated
    $workItemsCompleted = $projectStatsJson.dataProviders.'ms.vss-work-web.work-item-metrics-data-provider-verticals'.workMetrics.workItemsCompleted
    $commitsPushed = $projectStatsJson.dataProviders.'ms.vss-code-web.code-metrics-data-provider-verticals'.gitmetrics.commitsPushedCount
    if (!$commitsPushed) {$commitsPushed = 0}

    $pullRequestsCreated = $projectStatsJson.dataProviders.'ms.vss-code-web.code-metrics-data-provider-verticals'.gitmetrics.pullRequestsCreatedCount
    if (!$pullRequestsCreated) {$pullRequestsCreated = 0}

    $pullRequestsCompleted = $projectStatsJson.dataProviders.'ms.vss-code-web.code-metrics-data-provider-verticals'.gitmetrics.pullRequestsCompletedCount
    if (!$pullRequestsCompleted) {$pullRequestsCompleted = 0}
           
    $uriBuildMetrics = $UriOrganization + "$($project.id)/_apis/build/Metrics/Daily?minMetricsTime=$($monthAgo)" 
    $buildMetricsResult = Invoke-RestMethod -Uri $uriBuildMetrics -Method get -Headers $AzureDevOpsAuthenicationHeader
    $totalBuilds = 0
    $buildMetricsResult.value | Where-Object {$_.name -eq 'TotalBuilds'} | ForEach-Object { $totalBuilds+= $_.intValue }

    $totalReleases = 0
   # $UriReleaseMetrics = $UriOrganizationRM + "$($project.id)/_apis/Release/metrics?minMetricsTime=minMetricsTime=$($monthAgo)"
   # $releaseMetricsResult = Invoke-RestMethod -Uri $UriReleaseMetrics -Method get -Headers $AzureDevOpsAuthenicationHeader
    #$releaseMetricsResult.value | ForEach-Object { $totalReleases+= $_.value }

   
   $obj = New-Object PSObject -Property @{
                        TeamProjectName = $project.name
                        TeamProjectCountWorkItemCreated = $workItemsCreated
                        TeamProjectCountWorkItemCompleted = $workItemsCompleted
                        TeamProjectCountCommitsPushed = $commitsPushed                     
                        TeamProjectCountPRsCreated = $pullRequestsCreated
                        TeamProjectCountPRsCompleted = $pullRequestsCompleted
                        TeamProjectCountBuilds = $totalBuilds                        
                        TeamProjectCountReleases = $totalReleases
                }

         $featurearr += $obj
    
}

$excel = New-Object -ComObject Excel.Application

# Hide the Excel application window
$excel.Visible = $false

# Create a new workbook
$workbook = $excel.Workbooks.Add()

# Get the first worksheet in the workbook
$worksheet = $workbook.Worksheets.Item(1)


# Define the column headers based on the object properties
$header = $featurearr[0].PSObject.Properties.Name

# Add the column headers to the worksheet
$row = 1
$col = 1
$header | ForEach-Object {
    $worksheet.Cells.Item($row, $col) = $_
    $col++
}

# Add the object array data to the worksheet
$row = 2
foreach ($obj in $featurearr) {
    $col = 1
    foreach ($prop in $obj.PSObject.Properties) {
        $worksheet.Cells.Item($row, $col) = $prop.Value
        $col++
    }
    $row++
}
# Get the range of cells in the row
$rowRange = $worksheet.Range("A1", "XFD1")

# Make the row bold
$rowRange.Font.Bold = $true
# Save the workbook as an Excel file
$workbook.SaveAs($filepath)

# Close the workbook and Excel application
$workbook.Close($true)
$excel.Quit()
