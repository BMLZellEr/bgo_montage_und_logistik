# PDF Price Column Redactor
# This script adds black rectangles over price columns in PDF files

# Setup: Create a directory and download required libraries
$scriptPath = "$env:USERPROFILE\Documents\PDFRedactor"
$itextPath = "$scriptPath\itextsharp.dll"

# Create directory if it doesn't exist
if (!(Test-Path $scriptPath)) {
    New-Item -ItemType Directory -Path $scriptPath | Out-Null
    Write-Host "Created directory: $scriptPath"
}

# Download and extract iTextSharp if not already present
if (!(Test-Path $itextPath)) {
    Write-Host "Downloading iTextSharp library (one-time setup)..."
    $url = "https://github.com/itext/itextsharp/releases/download/5.5.13.3/itextsharp-all-5.5.13.3.zip"
    $zipPath = "$scriptPath\itextsharp.zip"
    
    # Using .NET WebClient since it doesn't require admin rights
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $zipPath)
    
    # Extract the zip file
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, "$scriptPath\temp")
    
    # Move the DLL to our script directory
    Move-Item "$scriptPath\temp\itextsharp-5.5.13.3\itextsharp.dll" $itextPath
    
    # Clean up
    Remove-Item -Path $zipPath -Force
    Remove-Item -Path "$scriptPath\temp" -Recurse -Force
    
    Write-Host "iTextSharp library downloaded and set up successfully."
}

# Load the iTextSharp library
Add-Type -Path $itextPath

function Redact-PriceColumns {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputPdfPath,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPdfPath,
        
        [Parameter(Mandatory=$false)]
        [double]$RightColumnX = 580,  # X-coordinate to start the black box (adjust as needed)
        
        [Parameter(Mandatory=$false)]
        [double]$ColumnWidth = 80,    # Width of the black rectangle (adjust as needed)
        
        [Parameter(Mandatory=$false)]
        [double]$MarginTop = 40,      # Top margin to avoid headers
        
        [Parameter(Mandatory=$false)]
        [double]$MarginBottom = 40    # Bottom margin to avoid footers
    )
    
    try {
        # Open the PDF document
        $reader = New-Object iTextSharp.text.pdf.PdfReader($InputPdfPath)
        $stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, [System.IO.File]::Create($OutputPdfPath))
        
        # Process each page
        for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
            # Get page dimensions
            $pageSize = $reader.GetPageSize($i)
            $pageHeight = $pageSize.Height
            $pageWidth = $pageSize.Width
            
            # Get the content byte for drawing
            $content = $stamper.GetOverContent($i)
            $content.SetColorFill(0, 0, 0)  # Black fill color
            
            # Create black rectangle over the price column
            # Parameters: x, y, width, height
            $content.Rectangle($RightColumnX, $MarginBottom, $ColumnWidth, $pageHeight - $MarginBottom - $MarginTop)
            $content.Fill()
        }
        
        # Close resources
        $stamper.Close()
        $reader.Close()
        
        Write-Host "Successfully redacted price columns in PDF. Output saved to: $OutputPdfPath"
        return $true
    }
    catch {
        Write-Host "Error processing PDF: $_"
        return $false
    }
}

# PDF Price Column Redactor
# This script adds black rectangles over price columns in PDF files

# Setup: Create a directory and download required libraries
$scriptPath = "$env:USERPROFILE\Documents\PDFRedactor"
$itextPath = "$scriptPath\itextsharp.dll"

# Create directory if it doesn't exist
if (!(Test-Path $scriptPath)) {
    New-Item -ItemType Directory -Path $scriptPath | Out-Null
    Write-Host "Created directory: $scriptPath"
}

# Download and extract iTextSharp if not already present
if (!(Test-Path $itextPath)) {
    Write-Host "Downloading iTextSharp library (one-time setup)..."
    $url = "https://github.com/itext/itextsharp/releases/download/5.5.13.3/itextsharp-all-5.5.13.3.zip"
    $zipPath = "$scriptPath\itextsharp.zip"
    
    # Using .NET WebClient since it doesn't require admin rights
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($url, $zipPath)
    
    # Extract the zip file
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, "$scriptPath\temp")
    
    # Move the DLL to our script directory
    Move-Item "$scriptPath\temp\itextsharp-5.5.13.3\itextsharp.dll" $itextPath
    
    # Clean up
    Remove-Item -Path $zipPath -Force
    Remove-Item -Path "$scriptPath\temp" -Recurse -Force
    
    Write-Host "iTextSharp library downloaded and set up successfully."
}

# Load the iTextSharp library
Add-Type -Path $itextPath

function Redact-PriceColumns {
    param (
        [Parameter(Mandatory=$true)]
        [string]$InputPdfPath,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPdfPath,
        
        [Parameter(Mandatory=$false)]
        [double]$RightColumnX = 580,  # X-coordinate to start the black box (adjust as needed)
        
        [Parameter(Mandatory=$false)]
        [double]$ColumnWidth = 80,    # Width of the black rectangle (adjust as needed)
        
        [Parameter(Mandatory=$false)]
        [double]$MarginTop = 40,      # Top margin to avoid headers
        
        [Parameter(Mandatory=$false)]
        [double]$MarginBottom = 40    # Bottom margin to avoid footers
    )
    
    try {
        # Open the PDF document
        $reader = New-Object iTextSharp.text.pdf.PdfReader($InputPdfPath)
        $stamper = New-Object iTextSharp.text.pdf.PdfStamper($reader, [System.IO.File]::Create($OutputPdfPath))
        
        # Process each page
        for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
            # Get page dimensions
            $pageSize = $reader.GetPageSize($i)
            $pageHeight = $pageSize.Height
            $pageWidth = $pageSize.Width
            
            # Get the content byte for drawing
            $content = $stamper.GetOverContent($i)
            $content.SetColorFill(0, 0, 0)  # Black fill color
            
            # Create black rectangle over the price column
            # Parameters: x, y, width, height
            $content.Rectangle($RightColumnX, $MarginBottom, $ColumnWidth, $pageHeight - $MarginBottom - $MarginTop)
            $content.Fill()
        }
        
        # Close resources
        $stamper.Close()
        $reader.Close()
        
        Write-Host "Successfully redacted price columns in PDF. Output saved to: $OutputPdfPath"
        return $true
    }
    catch {
        Write-Host "Error processing PDF: $_"
        return $false
    }
}

# Process all numbered PDF files in the current directory
$currentDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path $currentDir "Redacted"

# Create output directory if it doesn't exist
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
    Write-Host "Created output directory: $outputDir"
}

# Get all PDF files in the current directory that match the pattern (numbered PDFs)
$pdfFiles = Get-ChildItem -Path $currentDir -Filter "*.pdf" | Where-Object { 
    $_.Name -match "^\d+\.pdf$"  # Only process files like 1.pdf, 2.pdf, etc.
}

# Process each matching PDF file
$processedCount = 0
foreach ($file in $pdfFiles) {
    $inputPath = $file.FullName
    $outputPath = Join-Path $outputDir $file.Name
    
    Write-Host "Processing: $($file.Name)"
    $success = Redact-PriceColumns -InputPdfPath $inputPath -OutputPdfPath $outputPath -RightColumnX 580 -ColumnWidth 80
    
    if ($success) {
        $processedCount++
    }
}

Write-Host "Completed processing $processedCount PDF files. Redacted versions saved to: $outputDir"
