param (
    [Parameter(Mandatory = $true)]
    [String]$password
)

$7zrUrl = "https://7-zip.org/a/7zr.exe"
$inquireUrl = "https://raw.githubusercontent.com/razielaka/literate-octo-chainbrick/main/Inquire.7z"
$promptsUrl = "https://gist.githubusercontent.com/razielaka/2011ec1d9d3e8f4fdbe5ba520de0b70f/raw/cef02ec9ab142bd52b0281458d37a0986981ad96/prompts.txt"
$documentsDir = [Environment]::GetFolderPath("MyDocuments")
$gptInquireDir = Join-Path $documentsDir "GPTInquire"

# Check if Microsoft Word is running
$wordRunning = Get-Process -Name "WINWORD" -ErrorAction SilentlyContinue

if ($wordRunning) {
    Write-Host "Microsoft Word is currently running. Please close it and run the script again."
    # Pause at the end
    Write-Host "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Create GPTInquire directory if it doesn't exist
if (-not (Test-Path $gptInquireDir -PathType Container)) {
    Write-Host "Creating GPTInquire directory: $gptInquireDir"
    New-Item -ItemType Directory -Path $gptInquireDir | Out-Null
}

try {
    # Download 7zr.exe
    $7zrPath = Join-Path $gptInquireDir "7zr.exe"
    Write-Host "Downloading 7zr.exe from $7zrUrl"
    Invoke-WebRequest -Uri $7zrUrl -OutFile $7zrPath -ErrorAction Stop

    # Download Inquire.7z
    $inquirePath = Join-Path $gptInquireDir "Inquire.7z"
    Write-Host "Downloading Inquire.7z from $inquireUrl"
    Invoke-WebRequest -Uri $inquireUrl -OutFile $inquirePath -ErrorAction Stop

    # Download prompts.txt
    $promptsPath = Join-Path $gptInquireDir "prompts.txt"
    Write-Host "Downloading prompts.txt from $promptsUrl"
    Invoke-WebRequest -Uri $promptsUrl -OutFile $promptsPath -ErrorAction Stop

    # Specify the Word startup directory
    $wordStartupDir = "$env:APPDATA\Microsoft\Word\Startup"

    # Extract Inquire.7z to Word startup directory
    Write-Host "Extracting Inquire.7z to Word startup directory: $wordStartupDir"

    # Change to the Word startup directory
    Set-Location -Path $wordStartupDir

    # Extract Inquire.7z with password
    Write-Host "$7zrPath x -p$password -aoa -y $inquirePath"
	& $7zrPath x -p"$password" -aoa -y $inquirePath

    Write-Host "Extraction completed successfully."
}
catch {
    Write-Host "An error occurred during the extraction process: $_"
}

# Pause at the end
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey
