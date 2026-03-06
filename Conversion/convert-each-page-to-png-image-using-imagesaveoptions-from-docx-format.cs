using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertDocxPagesToPng
{
    static void Main()
    {
        // Path to the source DOCX file.
        string docxPath = @"C:\Input\Document.docx";

        // Folder where the PNG images will be saved.
        string outputFolder = @"C:\Output\Pages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(docxPath);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Optional: set resolution or image size if needed.
        // options.Resolution = 300; // DPI
        // options.ImageSize = new System.Drawing.Size(1200, 1600); // pixels

        // Iterate through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to render only the current page (zero‑based index).
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name for the current page.
            string outFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as a PNG image.
            doc.Save(outFile, options);
        }
    }
}
