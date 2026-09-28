using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add content that spans multiple pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Ensure the document layout is up‑to‑date so that PageCount is accurate.
        doc.UpdatePageLayout();

        // Configure TIFF save options for a multipage output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use a PageSet that covers all pages of the document.
            PageSet = new PageSet(0, doc.PageCount - 1),

            // Use LZW compression as a safe default.
            TiffCompression = TiffCompression.Lzw
        };

        string outputPath = "output.tiff";

        // Save the document as a multipage TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("TIFF file was not created.");

        Console.WriteLine($"Document saved as multipage TIFF to '{Path.GetFullPath(outputPath)}'.");
    }
}
