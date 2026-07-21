using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for all output files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page Word document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3.");

        // Verify the document has the expected number of pages.
        if (doc.PageCount != 3)
            throw new InvalidOperationException("The sample document should contain 3 pages.");

        // -----------------------------------------------------------------
        // 2. Render the whole document to a single multi‑page TIFF file.
        //    For TIFF the default layout is MultiPageLayout.TiffFrames,
        //    but we set it explicitly to demonstrate the rule usage.
        // -----------------------------------------------------------------
        ImageSaveOptions multiPageTiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        multiPageTiffOptions.PageLayout = MultiPageLayout.TiffFrames();

        string multiPageTiffPath = Path.Combine(artifactsDir, "Multipage.tiff");
        doc.Save(multiPageTiffPath, multiPageTiffOptions);

        // Ensure the multi‑page TIFF was created.
        if (!File.Exists(multiPageTiffPath))
            throw new FileNotFoundException("Failed to create the multi‑page TIFF.", multiPageTiffPath);

        // -----------------------------------------------------------------
        // 3. Split the multi‑page TIFF into individual page images.
        //    We re‑render each page of the original document as a separate TIFF.
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            ImageSaveOptions singlePageOptions = new ImageSaveOptions(SaveFormat.Tiff);
            // PageSet expects a zero‑based page index.
            singlePageOptions.PageSet = new PageSet(pageIndex);

            string pageTiffPath = Path.Combine(artifactsDir, $"Page_{pageIndex + 1}.tiff");
            doc.Save(pageTiffPath, singlePageOptions);

            // Validate that the individual page image was saved.
            if (!File.Exists(pageTiffPath))
                throw new FileNotFoundException($"Failed to create page image {pageIndex + 1}.", pageTiffPath);
        }

        // -----------------------------------------------------------------
        // 4. Indicate successful completion.
        // -----------------------------------------------------------------
        Console.WriteLine("Multi‑page TIFF created and split into individual pages successfully.");
    }
}
