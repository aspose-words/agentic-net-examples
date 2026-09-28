using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with multiple pages.
        string sourceDocPath = "Sample.docx";
        CreateSampleDocument(sourceDocPath, 5);

        // Load the document.
        Document doc = new Document(sourceDocPath);

        // Determine the number of pages.
        int pageCount = doc.PageCount;
        if (pageCount == 0)
            throw new InvalidOperationException("The document contains no pages.");

        // Prepare output folder.
        string outputFolder = "Output";
        Directory.CreateDirectory(outputFolder);

        // Render each page to a separate TIFF file in parallel.
        Parallel.For(0, pageCount, i =>
        {
            // Aspose.Words uses zero‑based page indices.
            int pageIndex = i;

            // Clone the original document for thread‑safe processing.
            Document pageDoc = (Document)doc.Clone(true);

            var options = new ImageSaveOptions(SaveFormat.Tiff)
            {
                // Save only the current page (zero‑based index).
                PageSet = new PageSet(pageIndex)
            };

            string pageFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.tiff");
            pageDoc.Save(pageFile, options);
        });

        // Verify that all page TIFF files were created.
        var pageFiles = Directory.GetFiles(outputFolder, "Page_*.tiff")
                                 .OrderBy(f => f)
                                 .ToArray();

        if (pageFiles.Length != pageCount)
            throw new InvalidOperationException("Not all page TIFF files were generated.");

        // Combine pages into a single multipage TIFF using Aspose.Words directly.
        string combinedTiffPath = Path.Combine(outputFolder, "Combined.tiff");
        var combinedOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // By default ImageSaveOptions saves all pages, producing a multipage TIFF.
        doc.Save(combinedTiffPath, combinedOptions);

        // Validate that the combined TIFF exists.
        if (!File.Exists(combinedTiffPath))
            throw new InvalidOperationException("Failed to create the combined multipage TIFF.");

        // Optional: clean up individual page files.
        foreach (var file in pageFiles)
        {
            File.Delete(file);
        }

        // Indicate successful completion.
        Console.WriteLine($"Multipage TIFF created at: {Path.GetFullPath(combinedTiffPath)}");
    }

    // Helper method to create a DOCX with the specified number of pages.
    private static void CreateSampleDocument(string path, int pages)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 1; i <= pages; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < pages)
                builder.InsertBreak(BreakType.PageBreak);
        }
        doc.Save(path);
    }
}
