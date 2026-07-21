using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string baseDir = Directory.GetCurrentDirectory();
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample DOC file if none exist
        string sampleDocPath = Path.Combine(baseDir, "Sample.doc");
        if (!File.Exists(sampleDocPath))
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("Sample page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Sample page 2.");
            sampleDoc.Save(sampleDocPath);
        }

        // Iterate over all .doc files in the base directory
        string[] docFiles = Directory.GetFiles(baseDir, "*.doc");
        foreach (string docFile in docFiles)
        {
            // Load the document
            Document doc = new Document(docFile);

            // Configure image save options for TIFF with 250 DPI
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 250 // Sets both horizontal and vertical DPI
            };

            // Determine output TIFF path
            string tiffPath = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(docFile) + ".tiff");

            // Save the document as a multi‑page TIFF
            doc.Save(tiffPath, options);

            // Verify that the TIFF file was created
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF: {tiffPath}");
        }
    }
}
