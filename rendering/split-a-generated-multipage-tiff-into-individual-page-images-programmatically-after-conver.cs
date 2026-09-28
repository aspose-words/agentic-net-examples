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
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Prepare an output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Render the whole document to a multipage TIFF.
        string multiTiffPath = Path.Combine(outputDir, "multipage.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // No PageSet assigned means all pages will be rendered.
        doc.Save(multiTiffPath, tiffOptions);

        // Ensure the multipage TIFF was created.
        if (!File.Exists(multiTiffPath))
            throw new InvalidOperationException("Multipage TIFF was not created.");

        // Split the multipage TIFF into individual page TIFF files.
        int pageCount = doc.PageCount;
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            string pageTiffPath = Path.Combine(outputDir, $"page_{pageIndex + 1}.tiff");
            ImageSaveOptions pageOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };
            doc.Save(pageTiffPath, pageOptions);

            // Verify each page image exists.
            if (!File.Exists(pageTiffPath))
                throw new InvalidOperationException($"TIFF for page {pageIndex + 1} was not created.");
        }
    }
}
