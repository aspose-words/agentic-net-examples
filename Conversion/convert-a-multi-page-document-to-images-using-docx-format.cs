using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Drawing;

class ConvertDocxToImages
{
    static void Main()
    {
        // Path to the source DOCX file (multi‑page document)
        string inputPath = @"C:\Docs\MultiPageDocument.docx";

        // Directory where the page images will be saved
        string outputDir = @"C:\Docs\Images";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputDir);

        // Load the DOCX document
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for PNG format
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Optional: adjust resolution or image size if required
        // options.Resolution = 300;
        // options.ImageSize = new Size(1240, 1754); // Example size

        // Iterate through each page of the document
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Specify which page to render
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name for the current page
            string outputPath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.png");

            // Save the selected page as an image
            doc.Save(outputPath, options);
        }
    }
}
