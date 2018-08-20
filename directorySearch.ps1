#Variables
$date = Get-Date
$date = $date.ToString("yyyy-MM-dd")
$fromDirectory = "H:\Old" # Change this, directory to search for directories older than 365 days.
$toDirectory = "H:\Test" # Change this, directory where you want to move the directories older than 365 days.
$toDirectoryZip = "H:\Test\backup.$date.zip" # Change this, directory where you want to store backups for data retention incase something happens.
$errorFile = "H:\Test\error.txt" # Change this, file where errors can be written.

function Move-Directory {
    $directoryList = Get-ChildItem -Directory $fromDirectory
    foreach ($directory in $directoryList) {
        if ($directory.LastWriteTime -lt (Get-Date).AddDays(-365)) {
            Echo "Moving the following $directory from $fromDirectory to $toDirectory"
            Move-Item -Path $fromDirectory\$directory -Destination $toDirectory -WhatIf -ErrorAction Stop # Remove -WhatIf to actually run it
        }
    }
}

# Possible method for preventing the loss of data, create a zip file first then execute the move. 
# This zips the entire directory, not just items over 365 days.
function Create-Zip {
    # if a zip of the same name exists, remove it before creating another zip
    if (Test-Path $toDirectoryZip) {
        Remove-Item $toDirectoryZip
    }
    Add-Type -assembly "system.io.compression.filesystem"
    [io.compression.zipfile]::CreateFromDirectory($fromDirectory, $toDirectoryZip)
}
try
{
    Create-Zip
    Move-Directory
}
catch
{
    Echo "Error occurred on $date : $_" | Add-Content $errorFile
}
