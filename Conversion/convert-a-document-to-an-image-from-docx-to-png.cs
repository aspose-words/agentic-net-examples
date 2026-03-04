using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Folder where the PNG images will be saved.
        string outputFolder = @"C:\Docs\Images\";

        // Ensure the output folder exists.
        System.IO.Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(sourcePath);

        // Create ImageSaveOptions for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set resolution (dpi) for higher quality.
            Resolution = 300,
            // Optional: render each page separately.
            PageSet = PageSet.All
        };

        // Save each page as a separate PNG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Set the page to render (zero‑based index).
            pngOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name.
            string outputPath = System.IO.Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as PNG.
            doc.Save(outputPath, pngOptions);
        }
    }
}
