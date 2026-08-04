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

        // -----------------------------------------------------------------
        // 1. Create a sample multi‑page document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Page 1
        builder.Writeln("This is page 1.");
        builder.InsertBreak(BreakType.PageBreak);

        // Page 2
        builder.Writeln("This is page 2.");
        builder.InsertBreak(BreakType.PageBreak);

        // Page 3
        builder.Writeln("This is page 3.");

        // -----------------------------------------------------------------
        // 2. Render the whole document to a multi‑frame TIFF file.
        //    The default ImageSaveOptions.PageLayout for TIFF is TiffFrames,
        //    which creates one frame per page.
        // -----------------------------------------------------------------
        string multiTiffPath = Path.Combine(outputDir, "Multipage.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(multiTiffPath, tiffOptions);

        // Verify that the multi‑page TIFF was created.
        if (!File.Exists(multiTiffPath))
            throw new FileNotFoundException("Failed to create the multi‑page TIFF.", multiTiffPath);

        // -----------------------------------------------------------------
        // 3. Split the document into individual page images.
        //    Instead of reading the TIFF frames (which would require System.Drawing),
        //    we render each page again using a PageSet that contains a single page.
        // -----------------------------------------------------------------
        for (int i = 0; i < doc.PageCount; i++)
        {
            ImageSaveOptions pageOptions = new ImageSaveOptions(SaveFormat.Tiff);
            pageOptions.PageSet = new PageSet(i); // Zero‑based page index.

            string pagePath = Path.Combine(outputDir, $"Page_{i + 1}.tiff");
            doc.Save(pagePath, pageOptions);

            // Verify that each page image was created.
            if (!File.Exists(pagePath))
                throw new FileNotFoundException($"Failed to create page image for page {i + 1}.", pagePath);
        }

        // All operations completed successfully.
        Console.WriteLine("Multi‑page TIFF created at:");
        Console.WriteLine(multiTiffPath);
        Console.WriteLine("Individual page images created in the same folder.");
    }
}
