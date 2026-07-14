using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a folder to hold sample DOCX files and the resulting TIFFs.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // Step 1: Bootstrap a few sample documents locally.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 2; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            builder.Writeln($"Sample document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2.");

            string docPath = Path.Combine(workDir, $"Sample{i}.docx");
            sampleDoc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // Step 2: Iterate over all DOC/DOCX files, render each to a TIFF
        //         with a DPI of 250, and save the TIFF.
        // -----------------------------------------------------------------
        string[] docFiles = Directory.GetFiles(workDir, "*.doc*");

        foreach (string docFile in docFiles)
        {
            // Load the source document.
            Document doc = new Document(docFile);

            // Configure image save options for TIFF output.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set the desired resolution (DPI).
                Resolution = 250
            };

            // Determine the output TIFF file name.
            string tiffPath = Path.Combine(
                workDir,
                Path.GetFileNameWithoutExtension(docFile) + ".tiff");

            // Save the document as a (multi‑page) TIFF.
            doc.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF: {tiffPath}");

            Console.WriteLine($"Converted '{Path.GetFileName(docFile)}' to TIFF at {tiffPath}");
        }

        Console.WriteLine("All documents have been processed.");
    }
}
