# Set variables
$organization = "your-organization"
$project = "your-project"
$repository = "auto"
$apiVersion = "6.1-preview.1"
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"
$outputFile = "C:\Projects\AutoChurn\auto.xlsx"
$responseFile = "C:\Projects\AutoChurn\apiresp.txt"
# Calculate date range (last year)
$startDate = Get-Date "04/01/2022"
$endDate = (Get-Date).ToString("yyyy-MM-dd")

# Set API endpoint
$uri = "https://yourcompanyurl/yourproject/_apis/git/repositories/$repository/pullrequests?searchCriteria.status=completed&searchCriteria.fromDate=$startDate&searchCriteria.toDate=$endDate&searchCriteria.status=completed&`$top=500000&api-version=$apiVersion"

# Create header with personal access token
$header = @{
    Authorization = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$pat")))"
}

# Call API and retrieve data
$response = Invoke-RestMethod -Uri $uri -Headers $header -Method Get
$featurearr = @()
foreach ($pr in $response.value) {
    $commitUri = $pr.lastMergeCommit.url+"/changes"
    $commitResponse = Invoke-RestMethod -Uri $commitUri -Headers $header -Method Get
    Write-Host $commitUri
   
     $obj = New-Object PSObject -Property @{
                        ID = $pr.pullRequestId
                        Title = $pr.title
                        Status = $pr.status
                       LinesEdited = $commitResponse.changeCounts.Edit
                       LinesAdded = $commitResponse.changeCounts.Add
                       LinesDeleted = $commitResponse.changeCounts.Delete
                       CreatedBy = $pr.createdBy.displayName
                       CreationDate = Get-Date $pr.creationDate -Format "yyyy-MM-dd HH:mm:ss"
                       ClosedDate = $pr.closedDate
                       MergeStatus = $pr.mergeStatus
                }
                $featurearr += $obj
}

# Step 4: Save the list to Excel
# Use a PowerShell module like ImportExcel or EPPlus to export the $featureList to an Excel file
# Load the Excel COM object

# Load the Excel COM object
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
$workbook.SaveAs("C:\Projects\AutoChurn\auto.xlsx")

# Close the workbook and Excel application
$workbook.Close($true)
$excel.Quit()

