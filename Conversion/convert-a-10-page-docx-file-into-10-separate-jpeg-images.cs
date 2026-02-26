using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToJpegPages
{
    static void Main()
    {
        // Path to the source DOCX file (10 pages).
        string sourceFile = @"C:\Docs\SourceDocument.docx";

        // Folder where the JPEG images will be saved.
        string outputFolder = @"C:\Docs\PageImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(sourceFile);

        // Loop through each page in the document.
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Create ImageSaveOptions for JPEG format.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

            // Render only the current page (zero‑based index).
            options.PageSet = new PageSet(i);

            // Build the output file name (Page_1.jpg, Page_2.jpg, ...).
            string outFile = Path.Combine(outputFolder, $"Page_{i + 1}.jpg");

            // Save the single page as a JPEG image.
            doc.Save(outFile, options);
        }

        Console.WriteLine("Conversion completed. Images saved to: " + outputFolder);
    }
}
