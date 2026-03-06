using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderFirstPageToPng
{
    static void Main()
    {
        // Path to the source DOCX document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Path where the rendered PNG image will be saved.
        string outputPath = @"C:\Docs\FirstPage.png";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Render only the first page (zero‑based index 0).
        options.PageSet = new PageSet(0);

        // Save the first page as a PNG image.
        doc.Save(outputPath, options);
    }
}
