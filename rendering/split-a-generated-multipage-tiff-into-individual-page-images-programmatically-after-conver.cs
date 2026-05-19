using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three pages of simple text.
        builder.Writeln("Page 1 – Hello World!");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 – Aspose.Words");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 – End of Document");

        // -----------------------------------------------------------------
        // 2. Render the whole document to a single multi‑frame TIFF file.
        // -----------------------------------------------------------------
        ImageSaveOptions multiPageTiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Use the TiffFrames layout so each page becomes a separate frame.
        multiPageTiffOptions.PageLayout = MultiPageLayout.TiffFrames();
        string multiPageTiffPath = Path.Combine(artifactsDir, "Multipage.tiff");
        doc.Save(multiPageTiffPath, multiPageTiffOptions);

        // Verify that the multi‑page TIFF was created.
        if (!File.Exists(multiPageTiffPath))
            throw new FileNotFoundException("Failed to create the multi‑page TIFF.", multiPageTiffPath);

        // -----------------------------------------------------------------
        // 3. Split the document into individual page images.
        //    (Each page is saved as its own TIFF file.)
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            ImageSaveOptions singlePageOptions = new ImageSaveOptions(SaveFormat.Tiff);
            // Render only the current page (zero‑based index).
            singlePageOptions.PageSet = new PageSet(pageIndex);
            // Optional: set a higher resolution for clearer images.
            singlePageOptions.Resolution = 300;

            string pageTiffPath = Path.Combine(artifactsDir, $"Page_{pageIndex + 1}.tiff");
            doc.Save(pageTiffPath, singlePageOptions);

            // Verify that each page image was saved.
            if (!File.Exists(pageTiffPath))
                throw new FileNotFoundException($"Failed to save page {pageIndex + 1} image.", pageTiffPath);
        }

        // All operations completed successfully.
        Console.WriteLine("Multi‑page TIFF created and split into individual page images.");
    }
}
