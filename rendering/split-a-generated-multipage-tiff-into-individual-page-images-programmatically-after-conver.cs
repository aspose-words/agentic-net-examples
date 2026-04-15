using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all generated files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages with simple text.
        builder.Writeln("Page 1 – Hello World!");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 – Aspose.Words");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 – End of Document");

        // -----------------------------------------------------------------
        // 2. Render the whole document to a multi‑frame TIFF file.
        // -----------------------------------------------------------------
        string multiPageTiffPath = Path.Combine(artifactsDir, "Multipage.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Explicitly use the layout that creates one frame per page.
            PageLayout = MultiPageLayout.TiffFrames()
        };
        doc.Save(multiPageTiffPath, tiffOptions);

        // Verify that the multi‑page TIFF was created.
        if (!File.Exists(multiPageTiffPath))
            throw new FileNotFoundException("Failed to create the multi‑page TIFF.", multiPageTiffPath);

        // -----------------------------------------------------------------
        // 3. Split the multi‑page TIFF into separate page images.
        //    (Re‑render each page individually.)
        // -----------------------------------------------------------------
        int pageCount = doc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            // Configure options to render only the current page.
            ImageSaveOptions pageOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // PageSet uses zero‑based page indices.
                PageSet = new PageSet(i),
                // Optional: set a higher resolution for clearer images.
                Resolution = 300
            };

            string pageTiffPath = Path.Combine(artifactsDir, $"Page_{i + 1}.tiff");
            doc.Save(pageTiffPath, pageOptions);

            // Verify that the individual page image was created.
            if (!File.Exists(pageTiffPath))
                throw new FileNotFoundException($"Failed to create TIFF for page {i + 1}.", pageTiffPath);
        }

        // All operations completed successfully.
        Console.WriteLine("Multi‑page TIFF created and split into individual pages.");
    }
}
