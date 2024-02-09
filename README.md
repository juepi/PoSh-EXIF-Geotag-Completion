# PoSh-EXIF-Geotag-Completion
PowerShell script which adds missing geotags to JPG photos. I've written this script as my Sony Alpha camera often loses the Bluetooth connection to the smartphone (or the Sony "Imaging Edge Mobile" app), resulting in many photos with missing geotags.  
The script takes a directory as parameter (optionally recursive), analyzes the EXIF geotags of all photos and tries to complete missing tags from different photos according to a configurable maximum time difference (defaults to 30 minutes).  
The best match (closest "time taken" timestamp to the photo with the missing geotag) wins and the geotag will be copied. After completing the geotag of a photo, the "modified" timestamp of the file will be reset to the "time taken" timestamp to avoid sorting issues with the updated photos.  
**Attention:** If you have photos from multiple sources (e.g. camera, smartphone etc) in the analyzed (sub-)directory/ies, make sure that all sources had a correct (identical) time configured! Different local time settings on the devices will lead to wrong geotags being copied due to the timestamp based matching algorithm!

## Requirements

This script requires [TagLib#](https://github.com/mono/taglib-sharp/). You may extract the precompiled DLL from the [nuget package](https://www.nuget.org/api/v2/package/TagLibSharp) and copy it to the `lib` subdirectory where the PoSh script resides.

## Usage

`Complete-GPS-EXIF-data.ps1 -InputPath C:\Dir\to\Photos [-Recurse]`

Self-explanatory i guess :wink:
  

Have fun,  
Juergen