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
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Optional: set resolution (dpi) if needed.
        // pngOptions.Resolution = 300;

        // Loop through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the PageSet to render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.png, Page_2.png, ...).
            string outFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as a PNG image.
            doc.Save(outFile, pngOptions);
        }

        Console.WriteLine("Conversion completed. PNG files are saved in: " + outputFolder);
    }
}
