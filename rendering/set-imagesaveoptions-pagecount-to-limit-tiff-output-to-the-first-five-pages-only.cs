using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with more than five pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure the source document has at least five pages.
        if (doc.PageCount < 5)
            throw new InvalidOperationException("Document does not contain enough pages.");

        // Configure ImageSaveOptions for TIFF and limit to the first five pages.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use PageSet to specify the range of pages to render (zero‑based start index).
            PageSet = new PageSet(0, 5) // Render pages 1‑5.
        };

        string outputPath = "output.tiff";
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("TIFF output was not created.", outputPath);
    }
}
