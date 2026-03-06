using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderPagesToTiff
{
    static void Main()
    {
        // Path to the source Word document.
        string inputPath = @"C:\Docs\Input.docx";

        // Path where the resulting TIFF will be saved.
        string outputPath = @"C:\Docs\Output.tiff";

        // Load the document (create/load rule).
        Document doc = new Document(inputPath);

        // Define the page range to render.
        // PageRange uses zero‑based indices: 0 = first page, 2 = third page.
        PageRange range = new PageRange(0, 2); // pages 1‑3

        // Create a PageSet from the range.
        PageSet pageSet = new PageSet(range);

        // Configure image save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageSet = pageSet;                     // Render only the specified pages.
        options.PageLayout = MultiPageLayout.TiffFrames(); // Each page as a separate frame.

        // Save the document using the configured options (save rule).
        doc.Save(outputPath, options);
    }
}
