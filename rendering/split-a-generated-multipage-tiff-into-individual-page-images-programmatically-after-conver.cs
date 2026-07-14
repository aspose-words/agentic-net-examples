using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1 – Hello World!");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 – Aspose.Words");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 – Split TIFF example.");

        // -----------------------------------------------------------------
        // 1. Render the whole document to a multi‑page TIFF file.
        // -----------------------------------------------------------------
        string multiPageTiffPath = Path.Combine(outputDir, "Document.MultiPage.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // TiffFrames layout renders each page as a separate frame in the TIFF.
            PageLayout = MultiPageLayout.TiffFrames()
        };
        doc.Save(multiPageTiffPath, tiffOptions);

        // Verify that the multi‑page TIFF was created.
        if (!File.Exists(multiPageTiffPath))
            throw new FileNotFoundException("Failed to create the multi‑page TIFF.", multiPageTiffPath);

        // -----------------------------------------------------------------
        // 2. Split the multi‑page TIFF into individual page images.
        //    (Render each page separately using PageSet.)
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            string pageTiffPath = Path.Combine(outputDir, $"Document.Page{pageIndex + 1}.tiff");

            ImageSaveOptions singlePageOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render only the current page (zero‑based index).
                PageSet = new PageSet(pageIndex)
            };

            doc.Save(pageTiffPath, singlePageOptions);

            // Verify that each split page image was created.
            if (!File.Exists(pageTiffPath))
                throw new FileNotFoundException($"Failed to create page image for page {pageIndex + 1}.", pageTiffPath);
        }

        // All operations completed successfully.
        Console.WriteLine("Multi‑page TIFF created and split into individual pages.");
    }
}
