using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentToTiff
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the resulting TIFF will be saved.
        string outputPath = @"C:\Docs\Output.tiff";

        // Load the document.
        Document doc = new Document(inputPath);

        // Define the page range to render.
        // PageRange uses zero‑based indices: this example renders pages 2‑4 (indices 1‑3).
        PageSet pageSet = new PageSet(new PageRange(1, 3));

        // Create ImageSaveOptions for TIFF format.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Apply the page range.
        tiffOptions.PageSet = pageSet;

        // Render each selected page as a separate frame in a multi‑frame TIFF.
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Save the document as a TIFF using the configured options.
        doc.Save(outputPath, tiffOptions);
    }
}
