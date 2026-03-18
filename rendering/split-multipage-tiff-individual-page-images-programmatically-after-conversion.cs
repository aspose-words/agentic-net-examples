using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffSplitter
{
    static void Main()
    {
        // Load the source document (any format supported by Aspose.Words).
        Document doc = new Document("input.docx");

        // Iterate through each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Create ImageSaveOptions for TIFF output.
            ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

            // Ensure each page is rendered as a separate frame in the TIFF.
            // For TIFF the default layout is TiffFrames, but we set it explicitly.
            tiffOptions.PageLayout = MultiPageLayout.TiffFrames();

            // Select only the current page to render.
            tiffOptions.PageSet = new PageSet(pageIndex);

            // Optional: set resolution or image size if required.
            // tiffOptions.Resolution = 300;
            // tiffOptions.ImageSize = new Size(2000, 2600);

            // Save the single‑page TIFF to a file.
            string outputPath = $"Page_{pageIndex + 1}.tiff";
            doc.Save(outputPath, tiffOptions);
        }

        Console.WriteLine("Document pages have been split into individual TIFF files.");
    }
}
