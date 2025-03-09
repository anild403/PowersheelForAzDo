$folderPath = "C:\Projects\TFVC\Source"

$fileCount = (Get-ChildItem -Path $folderPath -Recurse -File -Filter "*.txt").Count

Write-Host "The total number of .txt files in $folderPath and its subfolders is: $fileCount"
