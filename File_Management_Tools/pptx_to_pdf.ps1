#=================================================================================================[VARIABLES AND COLOUR VARIABLES]

$esc = "$([char]27)"

#Colours
$Border = "DarkGray"
$Heading = "Cyan"

$InfoText = "Gray"
$FilesText = "Yellow"
$StatusText = "DarkYellow"

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
Write-Host "This script's purpose is to find all .ppt and .pptx files in a" -ForegroundColor ${InfoText} -NoNewLine
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "source directory (and its subfolders) and convert them to PDF." -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "You will be asked for a source directory and a destination directory." -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "The script will recreate the source folder structure in the" -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host "|" -ForegroundColor ${Border} -NoNewline
Write-Host "destination folder to store the converted PDF files." -NoNewline -ForegroundColor ${InfoText}
Write-Host "${esc}[80G|" -ForegroundColor ${Border}

Write-Host ${HorizontalLine} -ForegroundColor ${Border}
Write-Host ""

#=======================================================================================================[SOURCE DIRECTORY PROMPT]

$sourcePath = $null
while ($true)
{
    Write-Host "Please enter the full path to the SOURCE directory (where your .ppt/.pptx files are)"
    $sourcePath = Read-Host "Source Directory Path"

    # Check if the user-provided path exists and is a directory
    if (Test-Path $sourcePath -PathType Container)
    {
        # Path is valid, break the loop
        break
    }
    else
    {
        Write-Host "Error: The path '$sourcePath' does not exist or is not a directory. Please try again." -ForegroundColor Red
        Write-Host "" # Add a blank line for readability
    }
}

Write-Host "" # Add a blank line
Write-Host "Source directory set to: $sourcePath" -ForegroundColor Green
Write-Host "" # Add a blank line

#==================================================================================================[DESTINATION DIRECTORY PROMPT]

$destinationPath = $null
while ($true)
{
    Write-Host "Please enter the full path to the DESTINATION directory (where the PDFs will be saved)"
    $destinationPath = Read-Host "Destination Directory Path"

    # Check if the user-provided path exists and is a directory
    if (Test-Path $destinationPath -PathType Container)
    {
        # Path is valid, break the loop
        break
    }
    else
    {
        # Ask if the user wants to create it
        Write-Host "The path '$destinationPath' does not exist." -ForegroundColor Yellow
        $choice = Read-Host "Do you want to create it? (Y/N)"
        if ($choice -eq 'y' -or $choice -eq 'Y')
        {
            try
            {
                New-Item -ItemType Directory -Path $destinationPath -Force -ErrorAction Stop | Out-Null
                Write-Host "Directory '$destinationPath' created successfully." -ForegroundColor Green
                break
            }
            catch
            {
                Write-Host "Error creating directory: $_" -ForegroundColor Red
                Write-Host "Please try again."
            }
        }
        else
        {
            Write-Host "Please enter a different destination path."
        }
    }
}

Write-Host "" # Add a blank line
Write-Host "Destination directory set to: $destinationPath" -ForegroundColor Green
Write-Host "" # Add a blank line

#================================================================================================================[FILE SCANNING]

Write-Host "Scanning for PowerPoint files..." -ForegroundColor $StatusText
$filesToConvert = Get-ChildItem -Path $sourcePath -Recurse -Include *.ppt, *.pptx

# Check if any files were found
if ($filesToConvert.Count -eq 0) 
{
    Write-Host "No .ppt or .pptx files found in '$sourcePath'." -ForegroundColor ${ErrorText}
    Write-Host ""
    Read-Host "Press Enter to exit..."
    exit
}

Write-Host "Found $($filesToConvert.Count) files to convert." -ForegroundColor $FilesText
Write-Host ${HorizontalLine} -ForegroundColor ${Border}
Write-Host ""

#==============================================================================================================[FILE CONVERSION]

# --- Create Y/N Prompt ---
$title = "Confirmation"
$message = "Do you want to start converting $($filesToConvert.Count) PowerPoint files?"

# Create the "Yes" choice
$yesChoice = [System.Management.Automation.Host.ChoiceDescription]::new(
    "&Yes", 
    "Select this to begin converting all found .ppt/.pptx files to PDF."
)

# Create the "No" choice
$noChoice = [System.Management.Automation.Host.ChoiceDescription]::new(
    "&No", 
    "Select this to cancel the operation and exit the script. No files will be changed."
)

$choices = [System.Management.Automation.Host.ChoiceDescription[]]@(
    $yesChoice,
    $noChoice
)

$defaultChoice = 1 # Default to No (index 1)

# --- Process the Choice ---
$result = $Host.UI.PromptForChoice($title, $message, $choices, $defaultChoice)

$numFiles = 0

switch ($result)
{
    0 { 
        # 0 is the index for "Yes"
        Write-Host "Starting Conversion..." -ForegroundColor $StatusText
        
        # Create a PowerPoint application object ONCE
        $ppt_app = $null
        try
        {
            $ppt_app = New-Object -ComObject PowerPoint.Application
            $saveFormat = [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF
            
            foreach ($file in $filesToConvert) 
            {
                # Get the relative path of the file from the source
                # .Substring($sourcePath.Length) ensures we get the path *after* the source folder
                $relativePath = $file.FullName.Substring($sourcePath.Length)
                
                # Create the full destination path, including the relative path
                $destinationFileFullName = Join-Path -Path $destinationPath -ChildPath $relativePath
                
                # Change the extension to .pdf
                $destinationFileFullName = [System.IO.Path]::ChangeExtension($destinationFileFullName, ".pdf")
                
                # Get the directory part of the destination path
                $destinationFileDirectory = [System.IO.Path]::GetDirectoryName($destinationFileFullName)
                
                # Create the destination directory if it doesn't exist
                if (-not (Test-Path $destinationFileDirectory))
                {
                    New-Item -ItemType Directory -Path $destinationFileDirectory -Force | Out-Null
                }
                
                Write-Host ${SeparationLine} -ForegroundColor ${LineBreak}
                Write-Host "Converting '$($file.Name)'..." -ForegroundColor ${StatusText}
                Write-Host "  Source: $($file.FullName)" -ForegroundColor $InfoText
                Write-Host "  Target: $destinationFileFullName" -ForegroundColor $InfoText
                
                # Open the presentation
                $presentation = $ppt_app.Presentations.Open($file.FullName, $true, $false, $false) # ReadOnly, Untitled, WithWindow
                
                # Save the presentation as PDF
                $presentation.SaveAs($destinationFileFullName, $saveFormat)
                
                # Close the presentation
                $presentation.Close()
                
                $numFiles++
            }
        }
        catch
        {
            Write-Host "An error occurred during conversion: $_" -ForegroundColor $ErrorText
        }
        finally
        {
            # Quit PowerPoint application
            if ($ppt_app -ne $null)
            {
                $ppt_app.Quit()
                
                # Clean up and release the COM object
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt_app) | Out-Null
            }
        }
        
        Write-Host ""
        Write-Host ${SeparationLine}
        Write-Host "CONVERSION COMPLETE. Number of files converted: " -ForegroundColor ${CompleteText} -NoNewline
        Write-Host "${numFiles}" -ForegroundColor ${FilesText}
        break
    } 
    1 { 
        # 1 is the index for "No"
        Write-Host "Cancelled Conversion." -ForegroundColor ${CancelledText}
        break
    }
}

# This replaces the 'PAUSE' command
Write-Host ""
Read-Host "Press Enter to exit..."