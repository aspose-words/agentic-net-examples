using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder to hold sample documents and the PowerShell script.
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "SampleDocs");
        Directory.CreateDirectory(baseDir);

        // Create a sample DOC file.
        string docPath = Path.Combine(baseDir, "Sample.doc");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for TIFF conversion.");
        doc.Save(docPath, SaveFormat.Doc);

        // Copy Aspose.Words.dll to the folder so the PowerShell script can load it.
        string assemblySource = typeof(Document).Assembly.Location;
        string assemblyDest = Path.Combine(baseDir, Path.GetFileName(assemblySource));
        File.Copy(assemblySource, assemblyDest, true);

        // Prepare the PowerShell script that converts DOC/DOCX files to TIFF with 250 DPI.
        string scriptPath = Path.Combine(baseDir, "ConvertDocsToTiff.ps1");
        string scriptContent = $@"
# PowerShell script to convert DOC/DOCX files to TIFF with 250 DPI using Aspose.Words
param(
    [string]$SourceFolder = '{baseDir}'
)

# Load Aspose.Words assembly (expects it in the same folder as the script)
$assemblyPath = Join-Path $PSScriptRoot '{Path.GetFileName(assemblyDest)}'
if (-not (Test-Path $assemblyPath)) {{
    Write-Error 'Aspose.Words.dll not found at ' + $assemblyPath
    exit 1
}}
Add-Type -Path $assemblyPath

# Get all .doc and .docx files recursively.
Get-ChildItem -Path $SourceFolder -Include *.doc, *.docx -Recurse | ForEach-Object {{
    $docFile = $_.FullName
    $tiffFile = [System.IO.Path]::ChangeExtension($docFile, '.tiff')
    $document = New-Object Aspose.Words.Document $docFile
    $options = New-Object Aspose.Words.Saving.ImageSaveOptions [Aspose.Words.SaveFormat]::Tiff
    $options.DpiX = 250
    $options.DpiY = 250
    $document.Save($tiffFile, $options)
    Write-Host ""Converted $docFile to $tiffFile""
}}
".Trim();

        // Write the script to disk.
        File.WriteAllText(scriptPath, scriptContent);

        // Verify that the script was created.
        if (!File.Exists(scriptPath))
        {
            throw new Exception("Failed to create the PowerShell script.");
        }

        // Program ends here; the generated script can be executed separately.
    }
}
