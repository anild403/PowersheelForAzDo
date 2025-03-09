#dotnet new uninstall project.ApiTemplate can be used

$templates = dotnet new --list

if($templates -like '*project-api-template*') {
      Write-Host '**project API Template is already Installed**'
} else {
      Write-Host '**project API Template is not Installed**'
      Write-Host '**Creating a new folder c:\Packages**'

      New-Item -ItemType Directory -Force C:\packages

      Write-Host '**Copying the nupkg file to c:\packages**'
      copy "\\Fsdc01\selectiv2\ISD\Claimwks\firstname\API Template\project.ApiTemplate.6.0.0.nupkg" "c:\packages\"

      Write-Host '**Installing the .net project template project-api-template**'
      dotnet new install C:\packages\project.ApiTemplate.6.0.0.nupkg


}

