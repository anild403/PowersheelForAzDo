# Set variables

$queryId = "8e79d270-4f97-4ce1-84b2-6ab6acdd9479"
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"

# Set up Azure DevOps connection
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat"))
#Query - Closed stories and features under epic
$uri = "https://yourcompanyurl/fa7ba97d-6ce3-4645-9325-9ad25bb6148e/_apis/wit/wiql/6ca29dde-2ffd-44f8-82e3-64f0fc4eacad?api-version=6.0"
$query = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Basic $base64AuthInfo"} -Method Get

Write-Host $query
$table = New-Object System.Collections.ArrayList
foreach ($item in $query.WorkitemRelations)
{
 # $uri = "https://yourcompanyurl/_apis/wit/workItems/" + ($item.target.id -join ",")
 $uri = $item.target.url
  $featureid = $item.target.id
  
  
  Write-Host $uri
  $workItems = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Basic $base64AuthInfo"} -Method Get
  Write-Host $workItems
  #$row =  $workItems | Select-Object id, fields.'System.Title', fields.'System.State', fields.'System.AssignedTo', fields.'System.Tags'
 $row= [PSCustomObject]@{ID=$workItems.id; Title=$workItems.fields.'System.Title';State=$workItems.fields.'System.State';Assigned=$workItems.fields.'System.AssignedTo';Tags=$workItems.fields.'System.Tags'}
  $table.Add($row)
# Write-Host $row
}

# Get work items from query
Write-Host $table



# Convert work items to table
#$table = $workItems.value | Select-Object id, fields.'System.Title', fields.'System.State', fields.'System.AssignedTo', fields.'System.Tags'

# Save table to Excel file
$table | Export-Excel -Path "C:\Projects\R4 - RemainingWork File.xlsx" -AutoSize -AutoFilter
