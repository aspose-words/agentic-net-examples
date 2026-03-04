using System;
using System.Drawing;               // For Color and Size structs
using Aspose.Words;                // Core document API
using Aspose.Words.Saving;         // ImageSaveOptions, SaveFormat, PageSet

class Program
{
    static void Main()
    {
        // Path to the source Word document.
        string inputFile = @"C:\Docs\SourceDocument.docx";

        // Path where the rendered image will be saved.
        string outputFile = @"C:\Docs\RenderedPage.png";

        // Load the document from the file system.
        Document doc = new Document(inputFile);

        // Create image save options for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Fine‑tuned rendering settings:
        // - Render at 300 DPI (higher quality than the default 96 DPI).
        options.Resolution = 300;

        // - Set the exact pixel dimensions of the output image (e.g., A4 at 300 DPI).
        options.ImageSize = new Size(2480, 3508);

        // - Enable anti‑aliasing for smoother edges.
        options.UseAntiAliasing = true;

        // - Use high‑quality rendering algorithms (slower but better results).
        options.UseHighQualityRendering = true;

        // - Render with a transparent background.
        options.PaperColor = Color.Transparent;

        // - If you need to render a specific page, set the PageSet property.
        //   Here we render only the first page (index 0).
        options.PageSet = new PageSet(0);

        // Save the document as an image using the configured options.
        doc.Save(outputFile, options);
    }
}
