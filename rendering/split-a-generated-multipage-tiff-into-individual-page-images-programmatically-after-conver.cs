using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 content.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 content.");

        // Ensure the layout is up‑to‑date so PageCount is accurate.
        doc.UpdatePageLayout();

        // Render the whole document to a multipage TIFF file.
        string multiPageTiffPath = "multipage.tiff";
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Save all pages.
        tiffOptions.PageSet = PageSet.All;
        doc.Save(multiPageTiffPath, tiffOptions);

        // Verify the multipage TIFF was created.
        if (!File.Exists(multiPageTiffPath))
            throw new FileNotFoundException("Failed to create the multipage TIFF.", multiPageTiffPath);

        // Split the multipage TIFF into individual page TIFF files.
        int pageCount = doc.PageCount; // Zero‑based count after layout update.
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            // Output file name uses 1‑based page numbers for readability.
            string singlePagePath = $"page_{pageIndex + 1}.tiff";

            ImageSaveOptions singlePageOptions = new ImageSaveOptions(SaveFormat.Tiff);
            // Render only the current page (zero‑based index).
            singlePageOptions.PageSet = new PageSet(pageIndex);
            doc.Save(singlePagePath, singlePageOptions);

            // Verify each single‑page TIFF was created.
            if (!File.Exists(singlePagePath))
                throw new FileNotFoundException($"Failed to create TIFF for page {pageIndex + 1}.", singlePagePath);
        }

        Console.WriteLine("Multipage TIFF and individual page TIFFs have been created.");
    }
}
