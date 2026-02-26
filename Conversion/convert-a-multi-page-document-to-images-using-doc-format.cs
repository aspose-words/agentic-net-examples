using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocToImages
{
    static void Main()
    {
        // Path to the source DOC file (multi‑page document).
        string sourceDocPath = @"C:\Input\MultiPageDocument.doc";

        // Directory where the page images will be saved.
        string outputImagesDir = @"C:\Output\PageImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputImagesDir);

        // Load the document from the file system.
        Document doc = new Document(sourceDocPath);

        // Create an ImageSaveOptions object for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Example: set resolution to 300 DPI.
            Resolution = 300,
            // Example: set image size (optional).
            ImageSize = new Size(1240, 1754) // A4 at 300 DPI.
        };

        // Iterate through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Render only the current page.
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name (page numbers are 1‑based for readability).
            string outputPath = Path.Combine(outputImagesDir, $"Page_{pageIndex + 1}.png");

            // Save the current page as an image.
            doc.Save(outputPath, options);
        }

        Console.WriteLine("Document pages have been converted to images successfully.");
    }
}
