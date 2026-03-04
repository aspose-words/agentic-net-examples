using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Create image save options for PNG output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Fine‑tune rendering quality.
        options.Resolution = 300;                     // 300 DPI.
        options.UseHighQualityRendering = true;       // Slow but high‑quality algorithms.
        options.UseAntiAliasing = true;               // Smooth edges.
        options.PaperColor = Color.Transparent;      // Transparent background.

        // Define a vertical multi‑page layout with 10 points spacing between pages.
        options.PageLayout = MultiPageLayout.Vertical(10);

        // Render each page to a separate PNG file.
        for (int i = 0; i < doc.PageCount; i++)
        {
            // Select the current page (zero‑based index) via PageSet.
            options.PageSet = new PageSet(i);

            // Save the page as an image.
            string outputPath = $"output_page_{i + 1}.png";
            doc.Save(outputPath, options);
        }

        // Optionally, render all pages into a single multi‑frame TIFF image.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.Resolution = 300;
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames(); // Each page becomes a TIFF frame.
        doc.Save("output_all_pages.tiff", tiffOptions);
    }
}
