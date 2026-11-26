using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add ten pages to the document.
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert a page break after each page except the last one.
            if (i < 10)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render only the first five pages using PageSet (start index 0, count 5).
            PageSet = new PageSet(0, 5)
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a multi‑page TIFF limited to five pages.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
    }
}
