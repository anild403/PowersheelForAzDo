
#URLs used
#https://yourcompanyurl/yourproject/_apis/tfvc/changesets/126294
#https://yourcompanyurl/_apis/tfvc/changesets/126294/changes
#https://yourcompanyurl/yourproject/_apis/tfvc/changesets/127160?includeDetails=true
#https://learn.microsoft.com/en-us/rest/api/azure/devops/tfvc/changesets/get?view=azure-devops-rest-7.1&tabs=HTTP

# Import the required PowerShell module
Install-Module -Name PowerShellGet -Force -AllowClobber
Install-Module -Name Az -AllowClobber -Force

# Set the Azure DevOps organization URL
$organizationUrl = "https://yourcompanyurl/"
$startdate = "09/01/2023"
$enddate = "12/12/2023"
# Set the project name
$projectName = "CommercialLines"

# Set the repository name
$repositoryName = "CommercialLines"
$pattern = "\d{7}"
# Set the branch name
$branchName = "CLAS_ST"

# Set the personal access token (PAT)
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"

# Set the path to save the Excel file
$excelFilePath = "C:\Projects\My Project\TFVC to Git\TFVCSTCommits.xlsx"

# Authenticate with Azure DevOps
Connect-AzAccount -AccessToken $pat
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)"))
$headers = @{Authorization=("Basic {0}" -f $base64AuthInfo)}
# Get the repository
#$repository = Get-AzReposRepository -Organization $organizationUrl -Project $projectName -Repository $repositoryName

# Get all the changesets to the specified branch using the Azure DevOps API
$uri = "https://yourcompanyurl/yourproject/_apis/tfvc/changesets?searchCriteria.itemPath=$/yourproject/Source/CLAS_ST&api-version=6.0&searchCriteria.fromDate=$startdate&searchCriteria.toDate=$enddate&`$skip=0&`$top=500000"
$changesetsResponse = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get

# Create an empty array to store the changeset details
$changesetDetails = @()

# Iterate through each changeset
foreach ($changeset in $changesetsResponse.value) {
    $changeseturi = "https://yourcompanyurl/yourproject/_apis/tfvc/changesets/$($changeset.changesetId)?includeDetails=true"
    $changesetRTSResponse = Invoke-RestMethod -Uri $changeseturi -Headers $headers -Method Get

    # Get the changeset details
    $changesetTitle = $changeset.comment
    $changesetAuthor = $changeset.author.displayName

    # Get the changeset details using the Azure DevOps API
    $changesetDetailsUri = "https://yourcompanyurl/_apis/tfvc/changesets/$($changeset.changesetId)/changes"
    $changesetDetailsResponse = Invoke-RestMethod -Uri $changesetDetailsUri -Headers $headers -Method Get

    # Iterate through each changed item in the changeset
    foreach ($change in $changesetDetailsResponse.value) {
        # Get the path of the changed item
        $changedFile = $change.item.path
        #$rts = ""
        #if ($changesetTitle -match $pattern) {
        #    # Extract the matched numbers into a variable
        #    $rts = $Matches[0]
        #
        #    # Print the matched numbers
        #    Write-Host "Matched Numbers: $rts"
        #} 
        $rts = $changesetRTSResponse.checkinNotes.value
        Write-Host "Matched Numbers: $($changesetRTSResponse.checkinNotes.value)"
        # Add the changeset details to the array
        $changesetDetails += [PSCustomObject]@{
            "Changeset Id" = $changeset.changesetId
            "Changeset Title" = $changesetTitle
            "Changed File" = $changedFile
            "Author Name" = $changesetAuthor
            "RTS" = $changesetRTSResponse.checkinNotes.value
            "Date" = $changeset.createdDate
            
        }
    }
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
$header = $changesetDetails[0].PSObject.Properties.Name

# Add the column headers to the worksheet
$row = 1
$col = 1
$header | ForEach-Object {
    $worksheet.Cells.Item($row, $col) = $_
    $col++
}

# Add the object array data to the worksheet
$row = 2
foreach ($obj in $changesetDetails) {
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
$workbook.SaveAs($excelFilePath)

# Close the workbook and Excel application
$workbook.Close($true)
$excel.Quit()


Write-Host "Changeset details have been saved to $excelFilePath."
