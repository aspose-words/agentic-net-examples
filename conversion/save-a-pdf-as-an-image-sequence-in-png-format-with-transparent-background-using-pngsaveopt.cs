using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing; // For Color

public class Program
{
    public static void Main()
    {
        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is page 3.");

        // Export each page to a separate PNG image with a transparent background.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Configure image save options.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.PaperColor = Color.Transparent; // Transparent background.
            options.PageSet = new PageSet(pageIndex); // Render only the current page.

            // Define output file name.
            string outputFileName = $"Page_{pageIndex + 1}.png";

            // Save the page as an image.
            doc.Save(outputFileName, options);

            // Validate that the image was created.
            if (!File.Exists(outputFileName))
                throw new InvalidOperationException($"Failed to create image file: {outputFileName}");
        }

        // All pages have been saved successfully.
        Console.WriteLine("PDF pages have been exported to PNG images with transparent backgrounds.");
    }
}
