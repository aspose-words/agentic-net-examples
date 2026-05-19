using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a working directory.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create sample DOCX files.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(workDir, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"Sample document {i} - Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Sample document {i} - Page 2.");
            doc.Save(docPath);
        }

        // Output folder for TIFF files.
        string tiffDir = Path.Combine(workDir, "TiffOutput");
        Directory.CreateDirectory(tiffDir);

        // Batch convert each DOCX to a multipage TIFF with custom DPI.
        foreach (string docFile in Directory.GetFiles(workDir, "*.docx"))
        {
            Document doc = new Document(docFile);

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Set desired DPI.
                Resolution = 300f,
                // Render all pages into a single multi‑frame TIFF.
                PageLayout = MultiPageLayout.TiffFrames()
            };

            string tiffPath = Path.Combine(tiffDir, Path.GetFileNameWithoutExtension(docFile) + ".tiff");
            doc.Save(tiffPath, options);

            // Verify that the file was created.
            if (!File.Exists(tiffPath))
                throw new InvalidOperationException($"Failed to create TIFF for {docFile}");
        }

        Console.WriteLine("Batch conversion to multipage TIFF completed successfully.");
    }
}
