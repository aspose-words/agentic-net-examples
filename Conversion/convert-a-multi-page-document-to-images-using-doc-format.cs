using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOC file (multi‑page document)
        string inputFile = "input.doc";

        // Directory where the page images will be saved
        string outputFolder = "PageImages";
        Directory.CreateDirectory(outputFolder);

        // Load the document from the file system
        Document doc = new Document(inputFile);

        // Configure image save options (JPEG format, 300 DPI)
        ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            Resolution = 300
        };

        // Render each page of the document to a separate image file
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to the current page (zero‑based index)
            imgOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name for the current page
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.jpg");

            // Save the current page as an image
            doc.Save(outputPath, imgOptions);
        }
    }
}
