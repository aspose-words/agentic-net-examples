using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Folder where the PNG files will be saved.
        string outputFolder = @"C:\Docs\Pages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Prepare image save options for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Iterate through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to render only the current page (zero‑based index).
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name for the current page.
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as a PNG image.
            doc.Save(outputPath, options);
        }
    }
}
