using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the multi‑page DOCX document.
        Document doc = new Document("input.docx");

        // Create the output folder if it does not exist.
        string outputFolder = "output";
        Directory.CreateDirectory(outputFolder);

        // Configure PNG rendering options.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Example: set a higher resolution for better quality.
        pngOptions.Resolution = 300;

        // Render each page of the document to a separate PNG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Restrict rendering to the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Build the file name for the current page.
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the page as a PNG image.
            doc.Save(outputPath, pngOptions);
        }
    }
}
