using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ConvertWordPageToPng
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path where the PNG image will be saved.
        string outputPath = @"C:\Images\Page1.png";

        // Load the Word document from file.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Render only the first page (zero‑based index = 0).
        options.PageSet = new PageSet(0);

        // Save the selected page as a PNG image.
        doc.Save(outputPath, options);
    }
}
