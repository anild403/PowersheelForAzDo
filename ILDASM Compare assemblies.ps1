# Create "Dumps" directories if they don't exist
$sourceDumpsPath = "C:\Source\Dumps"
$destinationDumpsPath = "C:\Destination\Dumps"
New-Item -ItemType Directory -Force -Path $sourceDumpsPath
New-Item -ItemType Directory -Force -Path $destinationDumpsPath

Get-ChildItem $sourceDumpsPath -Recurse | foreach ($_) {
    if ($_.PsIsContainer) {
        Remove-Item $_.FullName -Recurse
    } else {
        Remove-Item $_.FullName
    }
}
Get-ChildItem $destinationDumpsPath -Recurse | foreach ($_) {
    if ($_.PsIsContainer) {
        Remove-Item $_.FullName -Recurse
    } else {
        Remove-Item $_.FullName
    }
}
# Decompile each DLL in "C:\Source" and save dump in "C:\Source\Dumps"
Get-ChildItem -Path "C:\Source" -Filter *.dll | ForEach-Object {
    $dllName = $_.Name
    $dumpFilePath = Join-Path -Path $sourceDumpsPath -ChildPath "$dllName.txt"
    & "C:\Program Files (x86)\Microsoft SDKs\Windows\v194.0A\bin\NETFX 4.8 Tools\ildasm.exe" $_.FullName /output:$dumpFilePath
}

# Decompile each DLL in "C:\Destination" and save dump in "C:\Destination\Dumps"
Get-ChildItem -Path "C:\Destination" -Filter *.dll | ForEach-Object {
    $dllName = $_.Name
    $dumpFilePath = Join-Path -Path $destinationDumpsPath -ChildPath "$dllName.txt"
    & "C:\Program Files (x86)\Microsoft SDKs\Windows\v194.0A\bin\NETFX 4.8 Tools\ildasm.exe" $_.FullName /output:$dumpFilePath
}

# Compare contents of files in "C:\Source\Dumps" with files in "C:\Destination\Dumps"
Write-Host $sourceDumpsPath

Get-ChildItem $sourceDumpsPath -Filter *.txt | ForEach-Object {
    # Perform an action on each file
    Write-Host "Processing file: $($_.FullName), $($_.Name), "
    
    $destinationFile = Get-ChildItem -Path $destinationDumpsPath -Filter $($_.Name)
    if ($destinationFile) {
Write-Host " Destination File $($destinationFile.Name)."

        $sourceContent = Get-Content $($_.FullName)
        $destinationContent = Get-Content $destinationFile.FullName
        if ($sourceContent -ne $destinationContent) {
            Write-Host "File $($file.Name) has differences."
        }else {
        Write-Host "File$($file.Name) has no differences."
    }
    } else {
        Write-Host "File $($file.Name) does not exist in the destination folder."
    }
}


