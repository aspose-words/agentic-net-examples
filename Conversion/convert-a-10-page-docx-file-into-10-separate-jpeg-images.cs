using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file (10 pages).
        string sourcePath = @"C:\Docs\input.docx";

        // Folder where the JPEG images will be saved.
        string outputFolder = @"C:\Docs\OutputImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document using the provided Document constructor (load rule).
        Document doc = new Document(sourcePath);

        // Prepare the image save options for JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Loop through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to the current zero‑based page index (render one page only).
            jpegOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.jpg, Page_2.jpg, ...).
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.jpg");

            // Save the single page as a JPEG image using the Document.Save method (save rule).
            doc.Save(outputPath, jpegOptions);
        }

        Console.WriteLine("Conversion completed. JPEG images are saved in: " + outputFolder);
    }
}
