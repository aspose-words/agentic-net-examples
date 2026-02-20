using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentToPngConverter
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Directory where the PNG images will be saved.
        string outputDir = @"C:\Docs\PagesAsPng";
        Directory.CreateDirectory(outputDir);

        // Load the document from the DOC file.
        Document doc = new Document(inputPath);

        // Iterate through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Render only the current page (zero‑based index).
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.png, Page_2.png, ...).
            string outFile = Path.Combine(outputDir, $"Page_{pageIndex + 1}.png");

            // Save the selected page as a PNG image.
            doc.Save(outFile, options);
        }

        Console.WriteLine("All pages have been saved as PNG images.");
    }
}
