#get all the file extensions available in the directory
Get-Childitem -Recurse | Select-Object Extension -Unique



#get the lines of code
dir . -include ("*.cls", "*.frm", "*.bas") -Recurse -name | foreach{(GC $_).Count} | measure-object -sum