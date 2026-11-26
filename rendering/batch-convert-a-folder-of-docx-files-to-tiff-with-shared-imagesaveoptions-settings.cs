using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define base, source, and output folders.
        string baseFolder = Path.Combine(Path.GetTempPath(), "DocxToTiffDemo");
        string sourceFolder = Path.Combine(baseFolder, "SourceDocs");
        string outputFolder = Path.Combine(baseFolder, "TiffOutput");

        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files in the source folder.
        CreateSampleDocx(Path.Combine(sourceFolder, "Sample1.docx"),
                         "Hello World! This is the first document.");
        CreateSampleDocx(Path.Combine(sourceFolder, "Sample2.docx"),
                         "Second document with multiple paragraphs.\nParagraph two.\nParagraph three.");

        // Shared ImageSaveOptions for TIFF conversion.
        var tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300
            // Compression option omitted because the enum value is not available in this version.
        };

        // Perform batch conversion.
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");
        var generatedTiffs = new List<string>();

        foreach (string docxPath in docxFiles)
        {
            var doc = new Document(docxPath);
            string tiffPath = Path.Combine(outputFolder,
                                           Path.GetFileNameWithoutExtension(docxPath) + ".tiff");
            doc.Save(tiffPath, tiffOptions);
            generatedTiffs.Add(tiffPath);
        }

        // Validate that each TIFF file was created.
        foreach (string tiffPath in generatedTiffs)
        {
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"TIFF file was not created: {tiffPath}");
        }

        // Indicate successful conversion.
        Console.WriteLine($"Converted {generatedTiffs.Count} DOCX files to TIFF in folder: {outputFolder}");
    }

    // Helper method to create a simple DOCX file with given text.
    private static void CreateSampleDocx(string filePath, string content)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        string[] lines = content.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string line in lines)
        {
            builder.Writeln(line);
        }
        doc.Save(filePath);
    }
}
