<#

Author - Uday Modala
Date   - 9/2/2024
Description - This is the automated script for updating the latest version of the specified nuget package present in the given location with working project in visual studio.


Command to run -  powershell -File "C:\Users\pl1blabind\Desktop\uday\Updated_script.ps1" -fileLocation "C:\Users\pl1blabind\Desktop\uday\Config.json"


Requirements -
1)Provide the script location and config.json file location in the command to run.
2)Make sure that the Nuget client tools is installed in your system and add this location to environmental variables so that nuget command can be recognized every where.
3)Provide the outputDirectory in the config.json file as c:\ so that the new version will be available to the visual studio project.
4) msbuild path should be added in the environment variables

#>









param(
    [string]$fileLocation
)

$jsonContent = Get-Content -Raw -Path $fileLocation | ConvertFrom-Json

try{
    $packageName = $jsonContent.PackageName            #Name of the package that need to be updated
    $projectName = $jsonContent.ProjectName            #Name of the project in visual studio code
    $projectFilePath = $jsonContent.ProjectFilePath    #Project location 
    $sourceName = $jsonContent.SourceName              #Name of the source to be stored in Nuget source
    $sourceDestination=$jsonContent.SourceDestination  #Source where the nuget packages are available in Artifactory
    $userName= $jsonContent.UserName
    $password= $jsonContent.Password
    $apiKey= $jsonContent.ApiKey  
    $extension = $jsonContent.Extension    
    $outputDirectory = $jsonContent.OutputDirectory                                                                                                                                                      
}
catch{
    # Catching and handling any errors that might occur
    $errorMessage = $_.Exception.Message
    LogWrite "Exception Occured"
    LogWrite "Error: $errorMessage"
    Write-Host "Exception Occured"
    Write-Host "Error: $errorMessage"
    Write-Host "Done..." -ForegroundColor Green
    Exit
}



#Function to write information into the logfile
Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}


#Function to add source to the nuget sources.
function Add-NuGetSource {
    param (
        [string]$sourceName,
        [string]$sourceDestination,
        [string]$userName,
        [string]$password
    )

    # Run the nuget sources Add command
    nuget sources Add -Name $sourceName -Source $sourceDestination -username $userName -password $password
}


#Function to set the api key for the given source
function Set-NuGetApiKey {
    param (
        [string]$apiKey,
        [string]$sourceName
    )

    # Run the nuget setapikey command
    nuget setapikey $apiKey -Source $sourceName
}



#Function to get the list of versions available for ibex-driver-P
function Get-VersionsOfPackage {
    param (
        [string]$source,
        [string]$packageName,
        [string]$extension
    )

    $allVersions = nuget list -Source $sourceName   -AllVersions -PreRelease | 
        Where-Object {  $_ -like "$packageName *" } |
        ForEach-Object { $_.Replace($packageName, '') -replace [regex]::Escape($extension), '' }

    return $allVersions
}



#Function to fetch the latest version for any package
function Get-LatestVersion{
    param (
       [string[]]$allVersions
    )
    $versionObjects = $allVersions | ForEach-Object { [Version]$_ }
    $latestVersion = ($versionObjects | Sort-Object)[-1]
    return $latestVersion
}



#Function to get the latest version of specified series
function Get-LatestVersionInSeries {
        param (
            [string]$series,
            [string[]]$allVersions
        )

        # Filter versions that match the specified series
        $filteredVersions = $allVersions | Where-Object { $_ -like "$series.*" }

        # If there are matching versions, find the latest one
        if ($filteredVersions.Count -gt 0) {
            $latestVersion = $filteredVersions | Sort-Object { [Version]$_ } | Select-Object -Last 1
            LogWrite "Latest version in series $series is $latestVersion"
            return $latestVersion
        } else {
            LogWrite "No versions found in series $series."
        }
        return 
    }



#Function to get the version that is linked in visual studio
function Get-NuGetPackageVersionInVisualStudio {
    param (
        [string]$projectFilePath,
        [string]$packageName
    )

    # Check if the project file exists
    if (-not (Test-Path $projectFilePath)) {
        return 
    }

    # Load the content of the .csproj file as XML
    $projectXml = [xml](Get-Content $projectFilePath)

    # Check if the specified package is referenced in the project
    $packageReference = $projectXml.Project.ItemGroup.PackageReference | Where-Object { $_.Include -eq "$packageName" }
    $packageVersion = "Not"

    if ($packageReference -ne $null) {
        # Get the version of the referenced package
        $packageVersion = $packageReference.Version
        LogWrite "----The version of '$packageName' referenced in $projectFilePath is '$packageVersion'"
        Write-Host "----The version of '$packageName' referenced in $projectFilePath is '$packageVersion'"
    } else {
        LogWrite "----NuGet package '$packageName' is not referenced in $projectFilePath"
        Write-Host "----NuGet package '$packageName' is not referenced in $projectFilePath"
    }

    return $packageVersion
}



#Function to install the specified nuget package
function Install-NuGetPackage {
    param(
        [string]$PackageName,
        [string]$PackageVersion,
        [string]$sourceName,
        [string]$OutputDirectory
    )
    if (-not (Test-Path $OutputDirectory)) {
        New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
    }
    nuget install $PackageName -Version $PackageVersion -Source $sourceName -OutputDirectory $OutputDirectory
}



#Function to Remove the reference to the project
function Remove-NuGetPackage {
    param (
        [string]$projectFilePath,
        [string]$packageName,
        [string]$packageVersion
    )

    # Check if the project file exists
    if (-not (Test-Path $projectFilePath)) {
        Write-Host "Project file not found: $projectFilePath"
        return
    }

    # Load Project File as XML
    $projectXml = [xml](Get-Content $projectFilePath)

    # Find PackageReference to remove
    $packageReference = $projectXml.Project.ItemGroup.PackageReference | Where-Object {
        $_.Include -eq $packageName -and $_.Version -eq $packageVersion
    }

    if ($packageReference -ne $null) {
        # Remove the PackageReference from the ItemGroup
        $itemGroup = $packageReference.ParentNode
        $itemGroup.RemoveChild($packageReference) | Out-Null

        # Save Updated Project File
        $projectXml.Save($projectFilePath)

        # Log the Removal
        LogWrite "----NuGet package '$packageName' version '$packageVersion' removed from $projectFilePath"
    } else {
        LogWrite "----NuGet package '$packageName' version '$packageVersion' not found in $projectFilePath"
    }
}



#Function to add reference to the project
function Add-NuGetPackageVersion {
    param (
        [string]$projectFilePath,
        [string]$packageName,
        [string]$newPackageVersion
    )

    # Check if the project file exists
    if (-not (Test-Path $projectFilePath)) {
        LogWrite "Project file not found: $projectFilePath"
        return
    }

    # Load Project File as XML
    $projectXml = [xml](Get-Content $projectFilePath)

    # Find Existing ItemGroup containing PackageReference
    $itemGroup = $projectXml.Project.ItemGroup | Where-Object { $_.PackageReference -and $_.PackageReference.Include -eq $packageName }

    if ($itemGroup -eq $null) {
        # If no existing ItemGroup found, create a new one
        $itemGroup = $projectXml.CreateElement("ItemGroup")
        $projectXml.Project.AppendChild($itemGroup) | Out-Null
    }

    # Check if the PackageReference already exists
    $existingPackageReference = $itemGroup.PackageReference | Where-Object { $_.Include -eq $packageName -and $_.Version -eq $newPackageVersion }

    if ($existingPackageReference -eq $null) {
        # Create New PackageReference
        $newPackageReference = $projectXml.CreateElement("PackageReference")
        $newPackageReference.SetAttribute("Include", $packageName)
        $newPackageReference.SetAttribute("Version", $newPackageVersion)

        # Add New PackageReference to the existing or new ItemGroup
        $itemGroup.AppendChild($newPackageReference) | Out-Null

        # Save Updated Project File
        $projectXml.Save($projectFilePath)

        # Log the Update
        LogWrite "----NuGet package '$packageName' version '$newPackageVersion' added to $projectFilePath"
    } else {
        LogWrite "----NuGet package '$packageName' is already present in $projectFilePath with version '$newPackageVersion'"
    }
}


function Build-Project {
    param (
        [string]$projectFilePath
    )

    # Build the project using MSBuild
    $out = msbuild $projectFilePath /p:Configuration=Release
    $last = $out |Select-Object -Last 4
    LogWrite $last
    Write-Host $last

    # Check the exit code to determine success
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Build succeeded!"
        LogWrite "Build succeeded"
    } else {
        Write-Host "Build failed or did not meet the success criteria. Exit Code: $LASTEXITCODE"
        LogWrite "Build Failed."
        # Actions for a failed build
    }
}




#Execution starts from here

$sTime=Get-Date
# Get the current date and time in a different format
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
# Construct the log file name with the timestamp
$name=$packageName+"_"+$timestamp
$logFileName = "$name.log"
$tempDirectory = $env:TEMP
$LogFile = Join-Path -Path $tempDirectory -ChildPath $logFileName
Write-Host "Log file is generated at this location-$Logfile"


LogWrite "Time : $timestamp"
Add-Content -Path $Logfile -Value "`n"

LogWrite "***************Adding Artifactory source to Nuget sources****************"
try{
    if (nuget sources | Where-Object { $_ -like '*Artifactory*' }) {
        LogWrite "$sourceName source is already present."
    }
    else{
        $startTime = Get-Date
        Add-NuGetSource -sourceName $sourceName -sourceDestination $sourceDestination -userName $userName -password $password
        LogWrite "Successfully added the $sourceName to the Nuget sources."
        Write-Host "Successfully added the $sourceName to the Nuget sources."
        # Set ApiKey for the above source
        LogWrite "**************Setting api key for the source***********************"
        Write-Host "**************Setting api key for the source***********************"
        Set-NuGetApiKey -apiKey $apiKey -sourceName $sourceName
        LogWrite "ApiKey setup is done for the $sourceName source"
        Write-Host "ApiKey setup is done for the $sourceName source"
        LogWrite "*********************************************************************"
        Write-Host "*********************************************************************"
        Add-Content -Path $Logfile -Value "`n"
        $endTime= Get-Date
        $addSourceTime=$endTime-$startTime
        LogWrite "Time required to add the source to nuget: $addSourceTime"
        Write-Host "Time required to add the source to nuget: $addSourceTime"

    }
    LogWrite "**************************************************************************"
    Write-Host "**************************************************************************"
    Add-Content -Path $Logfile -Value "`n"
}
catch{
    # Catching and handling any errors that might occur
    $errorMessage = $_.Exception.Message
    LogWrite "Exception Occured"
    LogWrite "Error: $errorMessage"
    LogWrite "User Credentials provided by you might be incorrect."
    LogWrite "Please provide the correct details"
    Write-Host "Exception Occured"
    Write-Host "Error: $errorMessage"
    Write-Host "User Credentials provided by you might be incorrect."
    Write-Host "Please provide the correct details"
    Write-Host "Done..." -ForegroundColor Green
    Exit
}

try{
    $startTime = Get-Date
    $allVersions = Get-VersionsOfPackage -source $sourceName -packageName $packageName -extension $extension
    $endTime= Get-Date
    $listVersionsTime=$endTime-$startTime
    LogWrite "Time required to get list of versions: $listVersionsTime"
    Write-Host "Time required to get list of versions: $listVersionsTime"
    Add-Content -Path $Logfile -Value "`n"
}
catch{
    # Catching and handling any errors that might occur
    $errorMessage = $_.Exception.Message
    LogWrite "Exception Occured"
    LogWrite "Error: $errorMessage"
    LogWrite "Please check the source destination that you are provided."
    Write-Host "Exception Occured"
    Write-Host "Error: $errorMessage"
    Write-Host "Please check the source destination that you are provided."
    Write-Host "Done..." -ForegroundColor Green
    Exit
}


if($allVersions -eq $null){
    LogWrite "There is no package $packageName available in the given location."
    LogWrite "Please Specify the correct details."
    Write-Host "There is no package $packageName available in the given location."
    Write-Host "Please Specify the correct details."
    Write-Host "Done..." -ForegroundColor Green
    Exit
}


try{
    #Get the latest version of the specified package
    $startTime = Get-Date
    LogWrite "**********Getting the latest version from Artifactory*************"
    Write-Host "**********Getting the latest version from Artifactory*************"
    $latestVersion = Get-LatestVersion -allVersions $allVersions
    LogWrite "The latest version in Artfactory is $latestVersion ."
    Write-Host "The latest version in Artfactory is $latestVersion ."
    $endTime= Get-Date
    $latestVersionTime=$endTime-$startTime
    LogWrite "Time required to fetch latest version: $latestVersionTime"
    Write-Host "Time required to fetch latest version: $latestVersionTime"

    LogWrite "*******************************************************************"
    Write-Host "*******************************************************************"

    Add-Content -Path $Logfile -Value "`n"
}
catch{
    # Catching and handling any errors that might occur
    $errorMessage = $_.Exception.Message
    LogWrite "Exception Occured"
    LogWrite "Error: $errorMessage"
    LogWrite "The extension that was provided by you is incorrect."
    LogWrite "Please provide the correct details."
    Write-Host "Exception Occured"
    Write-Host "Error: $errorMessage"
    Write-Host "The extension that was provided by you is incorrect."
    Write-Host "Please provide the correct details."
    Write-Host "Done..." -ForegroundColor Green
    Exit
}



#Finding the version that is linked with the project in visual studio 
$startTime=Get-Date
LogWrite "********Finding the version that is linked in visual studio************"
Write-Host "********Finding the version that is linked in visual studio************"
$versionInVisual = Get-NuGetPackageVersionInVisualStudio -projectFilePath $projectFilePath -packageName $packageName
$endTime=Get-Date
if($versionInVisual -eq $null){
    LogWrite "Project file '$projectName.csproj' not found."
    LogWrite "The name of the project specified by you is incorrect."
    LogWrite "Please provide the correct Details."
    Write-Host "Project file '$projectName.csproj' not found."
    Write-Host "The name of the project specified by you is incorrect."
    Write-Host "Please provide the correct Details."
    Write-Host "Done..." -ForegroundColor Green
    Exit
}
$LinkedVersionTime=$endTime-$startTime
LogWrite "Time required to get the linked version in visual studio: $LinkedVersionTime"
LogWrite "*******************************************************************"
Write-Host "Time required to get the linked version in visual studio: $LinkedVersionTime"
Write-Host "*******************************************************************"
Add-Content -Path $Logfile -Value "`n"

try{
    $latest=$latestVersion.ToString() + $extension
    LogWrite "****Checking Whether the new version is linked with visual studio or not***"
    Write-Host "****Checking Whether the new version is linked with visual studio or not***"
    if($versionInVisual -eq $latest){
        LogWrite "$NewpackageVersion version is already installed and Linked to the project $projectName"
        Write-Host "$NewpackageVersion version is already installed and Linked to the project $projectName"
    }
    else{
        $LinkVersionTime = $null
        $startTime=Get-Date
        LogWrite "Installing the new version... "
        Write-Host "Installing the new version... "
        $result = Install-NuGetPackage -PackageName $packageName -PackageVersion $latest  -sourceName $sourceName -OutputDirectory $outputDirectory
        LogWrite $result
        Write-Host $result
        LogWrite "Installation Completed..."
        Write-Host "Installation Completed..."
        $endTime=Get-Date
        $installTime = $endTime-$startTime
        LogWrite "Time requried to install new version: $installTime"
        Write-Host "Time requried to install new version: $installTime"
        if($versionInVisual -eq "Not"){
            $startTime=Get-Date
            add-NuGetPackageVersion -projectFilePath $projectFilePath -packageName $packageName -newPackageVersion $latest
            $endTime=Get-Date
            $LinkVersionTime=$endTime-$startTime
            

        }
        else{
            $startTime=Get-Date
            add-NuGetPackageVersion -projectFilePath $projectFilePath -packageName $packageName -newPackageVersion $latest
            Remove-NuGetPackage -projectFilePath $projectFilePath -packageName $packageName -packageVersion $versionInVisual
            $endTime=Get-Date
            $LinkVersionTime=$endTime-$startTime
        }
        LogWrite "Time required to link new version in visual studio: $LinkVersionTime"
        Write-Host "Time required to link new version in visual studio: $LinkVersionTime"

    }
    LogWrite "*******************************************************************"
    Write-Host "*******************************************************************"
    Add-Content -Path $Logfile -Value "`n"
    
}
catch{
    # Catching and handling any errors that might occur
    $errorMessage = $_.Exception.Message
    LogWrite "Exception Occured"
    LogWrite "Error: $errorMessage"
    Write-Host "Exception Occured"
    Write-Host "Error: $errorMessage"
    Write-Host "Done..." -ForegroundColor Green
    Exit
}

Write-Host "Done..." -ForegroundColor Green


LogWrite "*******Build the $projectName in visual studio********"
Write-Host "*****Build the $projectName in visual studio********"
Build-Project -projectFilePath $projectFilePath
LogWrite "********************************************************"
Write-Host "******************************************************"
$eTime=Get-Date
$totalTime=$eTime -$sTime

LogWrite "Total Time taken: $totalTime"
Write-Host "Total Time taken: $totalTime"











