using System;
using Aspose.Words;
using Aspose.Words.Saving;

class WordToPngConverter
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path to the output PNG image.
        // This will contain the rendered page.
        string outputPath = @"C:\Docs\Sample_Page1.png";

        // Load the Word document from the file system.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Specify which page to render.
        // PageSet uses zero‑based page index, so 0 = first page.
        options.PageSet = new PageSet(0);

        // Optional: set resolution or background color if needed.
        // options.Resolution = 300;               // 300 DPI
        // options.PaperColor = System.Drawing.Color.Transparent;

        // Save the selected page as a PNG image.
        doc.Save(outputPath, options);
    }
}
