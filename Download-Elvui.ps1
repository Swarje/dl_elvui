$ExtractPath = "_retail_\Interface\AddOns\"
$DynPolicy = Get-ExecutionPolicy
$notfound = $False
$webPage = 'https://www.tukui.org'
$SourceUri = ($webPage + "/welcome.php")


Write-Host "Checking policy version and changinng to allow script to continue"
if ($DynPolicy -in @("restricted")){
    start-process  powershell.exe -Verb runas -ArgumentList "-noExit", "-command","Set-ExecutionPolicy bypass -Force"
}
#Get Drive letter
$drive = Read-Host "What drive does your WoW instance reside? Please include the letter, colon and slash  ex. C:\, D:\"
write-host "Depending on the size of your hard drive this can take up to a minute"

#Search the drive for wow folder
if ($Drive -match '[cdefg]:\\'){
    $quickPath = ($drive + "program files*")
    Write-Host 'Checking program files default paths'
    $wowpath = Get-ChildItem -path $quickPath -Recurse -Include "World of Warcraft" -ErrorAction SilentlyContinue{}
}

if(!($wowpath)){
    write-host 'Checking secondary pathways, this will take a bit longer'
    $wowpath = Get-ChildItem -path $drive -Recurse -Include "World of Warcraft" -ErrorAction SilentlyContinue{}
}

#Confirm viability of wowpath var
If ($wowpath -eq ("",$null)){
    $notfound = $True
}
if (!(test-path -Path $wowpath)){
    $notfound = $True
}
if ($notfound -eq $True){
    Write-Host "Wow cannot be found on your computer"
    Read-Host "Press enter to close"
    end
} 

$extractFullPath = join-path -Path $wowpath -ChildPath $ExtractPath
Write-Host ("Installation path is " + $extractFullPath)

# TODO: Check current elvui version against website version.


$ElvuiDownloadPath = ((Invoke-WebRequest $SourceUri).Links | ?{$_.innerHTML -like "Download Elvui*"}).href 
$DownloadSource = ($webPage + $ElvuiDownloadPath)
$tmpfolder = 'C:\Temp'
$ZipSource = ( $tmpfolder + '\' + $(Split-Path -Path $DownloadSource -Leaf))

#ensure temp exists
if (!(Test-Path $ZipSource)){
    New-Item -ItemType Directory $tmpfolder
}

#Download the file
Invoke-WebRequest -Uri $DownloadSource -OutFile $ZipSource

#Warn of location
write-host ("Extracting ElvUI files to " + $extractFullPath)

#Expand to wow folder
Expand-Archive -Path $ZipSource -DestinationPath $extractFullPath -Force

#TODO Confirm successful extraction.
if ($DynPolicy -notmatch "Get-ExecutionPolicy"){
    $Scriptblock = {
        $sb = {
                    write-host $args[0]
                    Set-ExecutionPolicy $args[0]
        }
    }
    write-host "Resetting execution policy as it was before Elvui Installation."
    start-process  powershell.exe -Verb runas -ArgumentList $Scriptblock, "& `$sb $DynPolicy" -Wait
    $finally = Get-ExecutionPolicy
    Write-Debug $finally
}

Read-Host "Press enter to close this window."
