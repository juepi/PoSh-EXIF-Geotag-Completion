# Requires TagLib# https://github.com/mono/taglib-sharp/
# https://www.nuget.org/api/v2/package/TagLibSharp

#region ###################### Script Parameters ##############################

[CmdletBinding()]
Param
(
    [Parameter(Mandatory = $true, Position = 1)]
    [ValidateScript({ if (Test-Path $_) { return $true } else { throw "InputPath does not exist." } })]
    [string]$InputPath,
    [Parameter(Mandatory = $false)]
    [switch]$Recurse
)
#endregion

#region ###################### External Libraries #############################

[System.Reflection.Assembly]::LoadFrom(("C:\Scripts\lib\TagLibSharp.dll")) | Out-Null

#endregion

#region ##################### User Settings ###################################

# Input Image file name extensions that will be handled
$SupportedInputFileExt = @('*.jpg', '*.jpeg')

# Maximum Time difference in minutes to accept GPS data for a Photo which is lacking GPS data
# NOTE: Match with smallest time difference wins!
$MaxTimeDiff = 30

#endregion

#region ##################### Global Vars #####################################
# Array with all photos which have valid GPS data
$GPS_Valid_Files = @()

# Array with all photos with missing GPS data
$GPS_Missing_Files = @()

# Info stored in Arrays for each photo (just for info!)
#$FileInfo = @{
#    FullName = ""
#    LastWriteTime = [datetime]""
#    Longitude = ""
#    Latitude = ""
#    Altitude = ""
#}

# Helpers
$BestMatchFile = ""
$BestMatchTimeDiff = ""


#region ############################## Main Script #####################################


Write-Host "Analyzing input files.." -ForegroundColor Yellow

# Analyze InputPath and extract files to handle
if (Test-Path $InputPath -PathType Container) {
    # InputPath is a directory, add * for get-childitem
    $InputPath = $($InputPath + '\*')
    # Handle every supported file in fodler (recursively, if requested)
    if ($Recurse) {
        $HandleItems = Get-ChildItem -Path $InputPath -Include $SupportedInputFileExt -Recurse
    }
    else {
        $HandleItems = Get-ChildItem -Path $InputPath -Include $SupportedInputFileExt
    }
    if ($HandleItems.Count -eq 0) {
        Write-Error "No Images to handle in path $InputPath!" -ErrorAction Stop
    }
}
else {
    # Input is a file, this is not supported
    Write-Error "InputPath must be a directory!" -ErrorAction Stop
}


$PhotoNumber = 1

$HandleItems | ForEach-Object {
    $Photo = $_
    # Show Progress Bar
    $Pct = [math]::Round((($PhotoNumber / $HandleItems.Count) * 100), 1)
    Write-Progress -Activity "Analyzing Photos.." -Status "$Pct% Complete:" -PercentComplete $Pct
    $WorkImg = [TagLib.File]::Create($Photo.FullName)
    if (-not $WorkImg.ImageTag.Longitude) {
        Write-Verbose "Photo $($Photo.Name) is missing Logitude Info."
        $FailFileInfo = [PSCustomObject]@{
            Name      = "$($Photo.Name)"
            FullName  = "$($Photo.FullName)"
            TimeTaken = [datetime]$WorkImg.ImageTag.DateTime
        }
        $GPS_Missing_Files += $FailFileInfo

    }
    else {
        Write-Verbose "$($Photo.Name) Longitude: [$($WorkImg.ImageTag.Longitude)]"
        $OKFileInfo = [PSCustomObject]@{
            Name      = "$($Photo.Name)"
            FullName  = "$($Photo.FullName)"
            TimeTaken = [datetime]$WorkImg.ImageTag.DateTime
            Longitude = $WorkImg.ImageTag.Longitude
            Latitude  = $WorkImg.ImageTag.Latitude
            Altitude  = $WorkImg.ImageTag.Altitude
        }
        $GPS_Valid_Files += $OKFileInfo
    }
    $PhotoNumber++
}


if ( $GPS_Missing_Files.Count -gt 0 ) {
    Write-Host "Got $($GPS_Missing_Files.Count) files with missing GPS data and $($GPS_Valid_Files.Count) with valid GPS data." -ForegroundColor Yellow
    if ($GPS_Valid_Files.Count -eq 0) {
        Write-Host "No valid GPS data sources, aborting." -ForegroundColor Red
        pause
        write-Error "No valid sources." -ErrorAction Stop
    }
    Write-Host "Trying to find matching GPS data and updating files.." -ForegroundColor Yellow

    $PhotoNumber = 1
    
    $GPS_Missing_Files | ForEach-Object {
        $FailPhoto = $_
        # Show Progress bar
        $Pct = [math]::Round((($PhotoNumber / $GPS_Missing_Files.Count) * 100), 1)
        Write-Progress -Activity "Handling $($FailPhoto.Name)" -Status "$Pct% Complete:" -PercentComplete $Pct
        Write-Host "Searching for suitable GPS data for $($FailPhoto.Name).." -ForegroundColor Magenta
        for ($i = 0; $i -lt $GPS_Valid_Files.length; $i++) {
            # Loop through valid data to find a timedate match
            $Diff = [Math]::Abs(($FailPhoto.TimeTaken - $GPS_Valid_Files[$i].TimeTaken).TotalMinutes)
            if ( $Diff -le $MaxTimeDiff ) {
                # Found a candidate
                if ($global:BestMatchTimeDiff) {
                    if ( $Diff -lt $global:BestMatchTimeDiff ) {
                        #Write-Host "BETTER Candidate found: $($Diff)" -ForegroundColor Green
                        $global:BestMatchFile = $i
                        $global:BestMatchTimeDiff = $Diff
                    }
                }
                else {
                    #Write-Host "FIRST Candidate found: $($Diff)" -ForegroundColor Yellow
                    $global:BestMatchFile = $i
                    $global:BestMatchTimeDiff = $Diff
                }
            }
            else {
                #Write-Host "UNSUITABLE candidate: $($Diff)" -ForegroundColor DarkYellow
            }
        }
        if ($global:BestMatchTimeDiff) {
            Write-Host "Best match Source photo: $($GPS_Valid_Files[$global:BestMatchFile].Name)" -ForegroundColor Green
            Write-Host "Best match time difference: $([math]::Round($global:BestMatchTimeDiff,1)) minutes" -ForegroundColor Green

            ############# Run EXIF Update Magic #####################
            Write-Host "Updating EXIF data.." -ForegroundColor Yellow
            $WorkImg = [TagLib.File]::Create($FailPhoto.FullName)
            #$GPS_Valid_Files[$global:BestMatchFile] | Format-Table *itude
            $WorkImg.ImageTag.Longitude = $GPS_Valid_Files[$global:BestMatchFile].Longitude
            $WorkImg.ImageTag.Latitude = $GPS_Valid_Files[$global:BestMatchFile].Latitude
            $WorkImg.ImageTag.Altitude = $GPS_Valid_Files[$global:BestMatchFile].Altitude
            Write-Host "Saving photo.." -ForegroundColor Yellow
            try { $WorkImg.Save() }
            catch { Write-Error "Failed to save file $($FailPhoto.Name)!" -ErrorAction SilentlyContinue ; pause }
            Write-Host "Correcting LastWriteTime.." -ForegroundColor Yellow
            Get-Item -Path $FailPhoto.FullName | ForEach-Object { $_.LastWriteTime = $FailPhoto.TimeTaken }
            Write-Host "Done." -ForegroundColor Green

        }
        else {
            Write-Host "NO candidate found for $($FailPhoto.Name)" -ForegroundColor Red
        }
        # Reset global vars for next file
        $global:BestMatchFile = ""
        $global:BestMatchTimeDiff = ""
        $PhotoNumber++
    }
}
else {
    Write-Host "No files with missing GPS data found in $($InputPath)" -ForegroundColor Green
}
# End
pause

#endregion