using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document (multi‑page Word file).
        Document doc = new Document("Input.docx");

        // Create image save options for TIFF format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages; you could also use new PageSet(0,1,2) for specific pages.
            PageSet = PageSet.All,

            // Set a high resolution for better quality.
            Resolution = 300,

            // Define a fixed pixel size for each rendered page.
            ImageSize = new Size(2000, 3000),

            // Use anti‑aliasing and high‑quality rendering algorithms.
            UseAntiAliasing = true,
            UseHighQualityRendering = true
        };

        // Configure a multi‑page layout: 2 columns per row with 20 pt spacing.
        options.PageLayout = MultiPageLayout.Grid(2, 20f, 20f);
        options.PageLayout.BackColor = Color.LightGray;   // background colour
        options.PageLayout.BorderColor = Color.Blue;      // page border colour
        options.PageLayout.BorderWidth = 2;               // border thickness

        // Save the document as a single multi‑page TIFF image.
        doc.Save("Output.tiff", options);
    }
}
