using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add content that spans multiple pages.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            // Insert enough lines to fill the page.
            for (int j = 0; j < 30; j++)
            {
                builder.Writeln($"Line {j + 1} on page {i}.");
            }

            // Add a page break after each page except the last.
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define the output TIFF file path.
        string outputPath = "output.tiff";

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use LZW compression for TIFF (compatible with all versions).
            TiffCompression = TiffCompression.Lzw,
            // Set the resolution (dpi) for the output images.
            Resolution = 300
            // No need to set PageSet; by default all pages are saved.
        };

        // Save the document as a multipage TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Output the number of pages in the source document.
        Console.WriteLine($"Document saved as multipage TIFF. Source page count: {doc.PageCount}");
    }
}
