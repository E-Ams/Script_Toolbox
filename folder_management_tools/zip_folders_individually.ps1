# --- Configuration ---
# Set the path to your 7-Zip executable
$7zipPath = "C:\Program Files\7-Zip\7z.exe"
# --- End Configuration ---

# Check if 7-Zip exists at the specified path
if (-not (Test-Path $7zipPath)) 
{
    Write-Host "Error: 7-Zip executable not found at '$7zipPath'" -ForegroundColor Red
    Write-Host "Please update the `$7zipPath variable in the script."
    Read-Host "Press Enter to exit..."
    exit
}

#=================================================================================================[VARIABLES AND COLOUR VARIABLES]

$esc = "$([char]27)"

#Colours


$Border = "DarkGray"
$Heading = "Cyan"

$InfoText = "Gray"
$FilesText = "Yellow"
$ZipText = "DarkYellow"

$CompleteText = "Green"
$CancelledText = "Blue"
$ErrorText = "Red"

$LineBreak = "DarkGray"

#Tables
$HorizontalLine = "+------------------------------------------------------------------------------+"
$SeparationLine = "--------------------------------------------------------------------------------"

#=============================================================================================================[INFO TABLE DISPLAY]

Write-Host ""
Write-Host ${HorizontalLine} -ForegroundColor ${Border}
Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "${esc}[34GSCRIPT INFO" -ForegroundColor ${Heading} -NoNewline
Write-Host "${esc}[80G|" -ForegroundColor ${Border}
Write-Host ${HorizontalLine} -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "This script's purpose is to find all the main sub-folders in a" -ForegroundColor ${InfoText} -NoNewLine
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "directory you specify and compress each one into a separate .zip file." -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "It uses 7-Zip (configured at the top of the script) to perform" -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "the compression." -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "The .zip files will be saved in the same directory you specified." -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host ${HorizontalLine} -ForegroundColor ${Border}
Write-Host ""

#===============================================================================================================[DIRECTORY PROMPT]

# --- Prompt User for Directory ---
$targetPath = $null
while ($true)
{
    Write-Host "Please enter the full path to the directory you want to scan for subfolders"
    Write-Host "to compress."
    $targetPath = Read-Host "Directory Path"

    # Check if the user-provided path exists and is a directory
    if (Test-Path $targetPath -PathType Container)
    {
        # Path is valid, break the loop
        break
    }
    else
    {
        Write-Host "Error: The path '$targetPath' does not exist or is not a directory. Please try again." -ForegroundColor Red
        Write-Host "" # Add a blank line for readability
    }
}

#==============================================================================================================[DIRECTORY DISPLAY]

Write-Host "" # Add a blank line
Write-Host "Target directory set to: $targetPath" -ForegroundColor Green

# --- End Prompt User ---

# Get all directories IN THE USER'S SPECIFIED FOLDER
$directories = Get-ChildItem -Path $targetPath -Directory

# Check if any directories were found
if ($directories.Count -eq 0) 
{
    Write-Host "No sub-directories found in '$targetPath' to compress." -ForegroundColor ${ErrorText}
    Write-Host ""
    Read-Host "Press Enter to exit..."
    exit
}

# List all directories that will be compressed
Write-Host ${HorizontalLine} -ForegroundColor ${Border}
Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "${esc}[24GCompress the following folders?" -NoNewline -ForegroundColor ${Heading} 
Write-Host "${esc}[80G|" -ForegroundColor ${Border}
Write-Host ${HorizontalLine} -ForegroundColor ${Border}

$lastIndex = $directories.Count - 1

for ($i = 0; $i -lt $directories.Count; $i++)
{
    $dir = $directories[$i]
    if ($i -eq $lastIndex)
    {
        # Use the 'corner' character for the last item (ASCII)
	Write-Host "|" -NoNewline -ForegroundColor ${Border}
        Write-Host " \-- $($dir.Name)" -NoNewline -ForegroundColor ${FilesText}
	Write-Host "${esc}[80G|" -ForegroundColor ${Border}
    }
    else
    {
        # Use the 'tee' character for all other items (ASCII)
	Write-Host "|" -NoNewline -ForegroundColor ${Border}
        Write-Host " |-- $($dir.Name)" -NoNewline -ForegroundColor ${FilesText}
	Write-Host "${esc}[80G|" -ForegroundColor ${Border}
    }
}

Write-Host ${HorizontalLine} -ForegroundColor ${Border}

#===============================================================================================================[FILE COMPRESSION]

# --- Create Y/N Prompt ---
$title = "Confirmation"
$message = "Do you want to start compression?"

# Create the "Yes" choice with its own help message
$yesChoice = [System.Management.Automation.Host.ChoiceDescription]::new(
    "&Yes", 
    "Select this to begin compressing all yellow-listed folders into .zip files."
)

# Create the "No" choice with its own help message
$noChoice = [System.Management.Automation.Host.ChoiceDescription]::new(
    "&No", 
    "Select this to cancel the operation and exit the script. No files will be changed."
)

# Create the "Info" choice with its own help message
$infoChoice = [System.Management.Automation.Host.ChoiceDescription]::new(
    "&Info", 
    "Select this to see a brief explanation of what the script does."
)

# Add the choice objects to the array
$choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    $yesChoice,
    $noChoice
)

$defaultChoice = 1 # Default to No (index 1)

# --- Process the Choice ---
$result = $Host.UI.PromptForChoice($title, $message, $choices, $defaultChoice)

$numFolders = 0;

switch ($result)
{
    0 { 
        # 0 is the index for "Yes"
        Write-Host "Starting Compression..."
            
        foreach ($dir in $directories) 
        {
            # Create the full path for the zip file in the *target* directory
            $archiveName = Join-Path -Path $targetPath -ChildPath "$($dir.Name).zip"

            # Use the full path for the target directory
            $targetDirectory = $dir.FullName
            Write-Host ${SeparationLine} -ForegroundColor ${LineBreak}
            Write-Host "Creating '$($dir.Name).zip'..." -ForegroundColor ${ZipText}
                
            # Use the Call Operator (&) to run the external .exe
            & $7zipPath a -tzip $archiveName $targetDirectory

	    $numFolders++
        }
        
	Write-Host ""
	Write-Host ${SeparationLine}
        Write-Host "COMPRESSION COMPLETE. Number of folders archived: " -ForegroundColor ${CompleteText} -NoNewline
        Write-Host "${numFolders}" -ForegroundColor ${FilesText}
        break # Exit the while loop
    } 
    1 { 
        # 1 is the index for "No"
        Write-Host "Cancelled Compression." -ForegroundColor ${CancelledText}
        break # Exit the while loop
    }
}

# This replaces the 'PAUSE' command
Write-Host ""
Read-Host "Press Enter to exit..."





