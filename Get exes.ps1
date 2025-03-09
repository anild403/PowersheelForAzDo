   $allfiles =  Get-ChildItem -path "D:\DOMAIN_NET\project\WebServices\ApplicationServices".Root -recurse -include "*.exe"

   foreach ($file in $allfiles)
   {
     Write-Host "name: " $file.name", date:"$file.LastWriteTime
   }
