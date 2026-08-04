using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define paths for the simulated network share, custom fonts folder, and output directory.
        string networkSharePath = Path.Combine(Path.GetTempPath(), "NetworkShare");
        string sourceDocPath = Path.Combine(networkSharePath, "Sample.docx");
        string fontsFolderPath = Path.Combine(Path.GetTempPath(), "CustomFonts");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "OutputTiffs");

        // Ensure required directories exist.
        Directory.CreateDirectory(networkSharePath);
        Directory.CreateDirectory(fontsFolderPath);
        Directory.CreateDirectory(outputDir);

        // Create a sample document if it does not already exist on the "network share".
        if (!File.Exists(sourceDocPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");
            sampleDoc.Save(sourceDocPath);
        }

        // Configure Aspose.Words to look for fonts in the custom fonts folder.
        // The folder may be empty; this demonstrates the API usage.
        FontSettings.DefaultInstance.SetFontsFolder(fontsFolderPath, recursive: true);

        // Load the document from the simulated network share.
        Document doc = new Document(sourceDocPath);

        // Prepare image save options for TIFF rendering.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300 // 300 DPI for decent quality.
        };

        // Render each page of the document to a separate TIFF file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            tiffOptions.PageSet = new PageSet(pageIndex);
            string outputPath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.tiff");
            doc.Save(outputPath, tiffOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {outputPath}");
        }

        // Optional: indicate successful completion.
        Console.WriteLine($"Rendered {doc.PageCount} page(s) to TIFF files in: {outputDir}");
    }
}
