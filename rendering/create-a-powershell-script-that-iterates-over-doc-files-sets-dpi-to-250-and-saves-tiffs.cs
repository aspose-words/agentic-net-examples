using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for temporary documents and output images.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a few sample DOCX files locally.
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln($"Sample document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2.");

            string docPath = Path.Combine(artifactsDir, $"Sample{i}.docx");
            sampleDoc.Save(docPath);
        }

        // Iterate over all DOC/DOCX files in the folder.
        foreach (string docFile in Directory.GetFiles(artifactsDir, "*.doc*"))
        {
            // Load the document.
            Document doc = new Document(docFile);

            // Configure image save options for TIFF with 250 DPI.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 250 // Sets both horizontal and vertical DPI.
            };

            // Determine output TIFF file name.
            string tiffPath = Path.ChangeExtension(docFile, ".tiff");

            // Save the entire document as a multipage TIFF.
            doc.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF file: {tiffPath}");
        }

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine("DOC files have been rendered to TIFF images at 250 DPI.");
    }
}
