using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document (optional, just to have a file).
        string docPath = Path.Combine(outputDir, "sample.docx");
        doc.Save(docPath);

        // Render the document to a multipage TIFF.
        string multiPageTiffPath = Path.Combine(outputDir, "multipage.tiff");
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(multiPageTiffPath, tiffOptions);

        // Verify the multipage TIFF was created.
        if (!File.Exists(multiPageTiffPath))
            throw new FileNotFoundException("Multipage TIFF was not created.", multiPageTiffPath);

        // Split the document pages into individual PNG images.
        int pageCount = doc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            // Configure PNG save options for a single page.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
            pngOptions.PageSet = new PageSet(i); // zero‑based page index

            string pageImagePath = Path.Combine(outputDir, $"page_{i + 1}.png");
            doc.Save(pageImagePath, pngOptions);

            // Verify the individual page image was saved.
            if (!File.Exists(pageImagePath))
                throw new FileNotFoundException($"Page image {i + 1} was not saved.", pageImagePath);
        }

        // All operations completed successfully.
        Console.WriteLine("Multipage TIFF split into individual page images.");
    }
}
