using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class MultiPageDocumentToImages
{
    static void Main()
    {
        // Path to the source document (any supported format, e.g., DOCX, PDF, etc.).
        string inputPath = @"C:\Input\MultiPageDocument.docx";

        // Folder where the resulting images will be saved.
        string outputFolder = @"C:\Output\Images";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document using the Document constructor (load rule).
        Document doc = new Document(inputPath);

        // Prepare image save options. Choose PNG format for lossless quality.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set resolution (dpi) if higher quality is needed.
            Resolution = 300
        };

        // Iterate through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to the current zero‑based page index.
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name, e.g., Page_1.png, Page_2.png, ...
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as an image (save rule).
            doc.Save(outputPath, options);
        }

        Console.WriteLine("Document pages have been successfully converted to images.");
    }
}
