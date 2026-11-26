using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define base, input, and output directories
        string baseDir = Path.Combine(Directory.GetCurrentDirectory(), "RenderDemo");
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputTiffs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a simple DOC file as sample input
        string sampleDocPath = Path.Combine(inputDir, "Sample.doc");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello Aspose.Words!");
        doc.Save(sampleDocPath, SaveFormat.Doc);

        // Path for the generated PowerShell script
        string scriptPath = Path.Combine(baseDir, "RenderToTiff.ps1");

        // PowerShell script that converts each DOC/DOCX in the input folder to a TIFF with 250 DPI
        string scriptContent = $@"
# PowerShell script to convert DOC/DOCX files to TIFF with 250 DPI using Aspose.Words
$assemblyPath = '{typeof(Document).Assembly.Location.Replace("\\", "\\\\")}'
Add-Type -Path $assemblyPath

$inputFolder = '{inputDir.Replace("\\", "\\\\")}'
$outputFolder = '{outputDir.Replace("\\", "\\\\")}'
if (-not (Test-Path $outputFolder)) {{
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}}

Get-ChildItem -Path $inputFolder -Filter *.doc* | ForEach-Object {{
    $docPath = $_.FullName
    $doc = [Aspose.Words.Document]::new($docPath)
    $options = New-Object Aspose.Words.Saving.ImageSaveOptions([Aspose.Words.SaveFormat]::Tiff)
    $options.DpiX = 250
    $options.DpiY = 250
    $outputPath = Join-Path $outputFolder ($_.BaseName + '.tiff')
    $doc.Save($outputPath, $options)
}}
".Trim();

        // Write the script to file
        File.WriteAllText(scriptPath, scriptContent);

        // Verify that the script was created
        if (!File.Exists(scriptPath))
        {
            throw new InvalidOperationException("Failed to create the PowerShell script.");
        }

        // Optional: Inform the user where the script is located
        Console.WriteLine($"PowerShell script created at: {scriptPath}");
    }
}
