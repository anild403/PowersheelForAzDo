# Set variables for Azure DevOps organization, project, feature ID and personal access token
$pat = "yzyzzyxzyxzxyyzzyzyyxyyyxxyxyxyzxzzyxzyzzzzzyxxyyzxz"
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($pat)"))
$headers = @{Authorization=("Basic {0}" -f $base64AuthInfo)}
# Step 1: Get all the features under a list of epics
$epics = @("469665", "482284", "484875")  # List of epics

$features = @()
$featurearr = @()
foreach ($epic in $epics) {
    # Use Azure DevOps REST API to retrieve features associated with the epic
    # Add the features to the $features array


    # Get the features associated with the epic
$url = "https://yourcompanyurl/_apis/wit/workItems?ids=$epic&`$expand=relations&api-version=6.0"
$response = Invoke-RestMethod -Uri $url -Headers $headers -Method Get

foreach($relation in $response.value.relations){
    
        $featureId = $relation.url.Split("/")[-1]
        #$featureUrl = "https://yourcompanyurl/_apis/wit/workItems/"+$featureId+"&`$expand=all?api-version=6.0"
        $featureUrl = "https://yourcompanyurl/_apis/wit/workItems?ids=$featureId&`$expand=relations&api-version=6.0"
        $feature = Invoke-RestMethod -Uri $featureUrl -Headers $headers -Method Get
        if($feature.value.fields.'System.WorkItemType' -eq "Feature"){

        $userStories = @()
        foreach($relation in $feature.value.relations){
            if($relation.rel -eq "System.LinkTypes.Hierarchy-Forward"){
                $userStoryId = $relation.url.Split("/")[-1]
                $userStoryUrl = $relation.url
                $userStory = Invoke-RestMethod -Uri $userStoryUrl -Headers $headers -Method Get
                if($userStory.fields.'System.WorkItemType' -eq "User Story"){
                    $userStories += $userStory
                }
            }
        }
        $storyPointsSum = 0
        foreach($userStory in $userStories){
            if($userStory.fields.'Microsoft.VSTS.Scheduling.StoryPoints'){
                $storyPointsSum += [int]$userStory.fields.'Microsoft.VSTS.Scheduling.StoryPoints'
            }
        }
        $list = $feature.value.fields.'System.Tags'
        $tshirtSizes = "XXS", "XS", "S", "M", "L", "XL", "XXL"

        # Split the string by semicolon
        $items = $list.Split(";")

        # Check if any of the items match with the T-shirt sizes
        $TShirtSize = ""
        foreach($item in $items){
            if($tshirtSizes -contains $item.Trim()){
                $TShirtSize = $item
                break
            }
        }
        $obj = New-Object PSObject -Property @{
                        ID = $feature.value.id
                        Title = $feature.value.fields.'System.Title'
                        WorkitemType = $feature.value.fields.'System.WorkItemType'
                        Parent = $feature.value.fields.'System.Parent'
                        Tags = $feature.value.fields.'System.Tags'
                        State = $feature.value.fields.'System.State'
                        StoryPointsSum = $storyPointsSum
                        TShirtSize = $TShirtSize
                }
         $featurearr += $obj
         $features += $feature
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
$workbook.SaveAs("C:\Projects\output.xlsx")

# Close the workbook and Excel application
$workbook.Close($true)
$excel.Quit()


