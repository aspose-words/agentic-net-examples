using System;
using System.IO;

public class Program
{
    public static void Main()
    {
        // PowerShell script that processes all .doc and .docx files in a folder,
        // renders each page to a TIFF image with 250 DPI, and saves the result.
        string scriptContent = @"
# Requires Aspose.Words for .NET assembly to be accessible.
param(
    [Parameter(Mandatory=$true)]
    [string]$SourceFolder,
    [Parameter(Mandatory=$true)]
    [string]$OutputFolder
)

# Load Aspose.Words assembly (adjust the path if necessary)
Add-Type -Path 'Aspose.Words.dll'

# Ensure output folder exists
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

# Get all DOC and DOCX files recursively
Get-ChildItem -Path $SourceFolder -Recurse -Include *.doc, *.docx | ForEach-Object {
    $docPath = $_.FullName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($docPath)
    $tiffPath = Join-Path $OutputFolder ($baseName + '.tiff')

    # Load the document
    $doc = New-Object Aspose.Words.Document($docPath)

    # Configure image save options for TIFF with 250 DPI
    $options = New-Object Aspose.Words.Saving.ImageSaveOptions([Aspose.Words.SaveFormat]::Tiff)
    $options.DpiX = 250
    $options.DpiY = 250

    # Save as multipage TIFF
    $doc.Save($tiffPath, $options)

    Write-Host ""Saved TIFF: $tiffPath""
}
";

        // Write the script to a .ps1 file in the current directory
        string scriptPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderDocsToTiff.ps1");
        File.WriteAllText(scriptPath, scriptContent);

        // Verify that the script file was created
        if (!File.Exists(scriptPath))
        {
            throw new InvalidOperationException("Failed to create the PowerShell script.");
        }

        // Optionally, inform the user (no interactive wait)
        Console.WriteLine($"PowerShell script generated at: {scriptPath}");
    }
}
