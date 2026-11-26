using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with multiple pages.
        const string docPath = "sample.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"This is page {i}.");
            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            if (i < 5)
                builder.InsertBreak(BreakType.PageBreak);
        }
        doc.Save(docPath);

        // Load the document and determine page count.
        var loadedDoc = new Document(docPath);
        loadedDoc.UpdatePageLayout();
        int pageCount = loadedDoc.PageCount;
        if (pageCount == 0)
            throw new Exception("Document contains no pages.");

        // Render the whole document to a multipage TIFF.
        const string outputTiffPath = "output.tiff";
        var saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages (from first to last) into a single multipage TIFF.
            PageSet = new PageSet(0, pageCount - 1),
            // Optional: set compression and resolution.
            TiffCompression = TiffCompression.Ccitt4,
            Resolution = 300
        };
        loadedDoc.Save(outputTiffPath, saveOptions);

        // Validate that the output TIFF exists and is non‑empty.
        if (!File.Exists(outputTiffPath))
            throw new Exception("Failed to create the output TIFF file.");

        var info = new FileInfo(outputTiffPath);
        if (info.Length == 0)
            throw new Exception("The generated TIFF file is empty.");

        // Indicate success.
        Console.WriteLine("Multipage TIFF generated successfully.");
    }
}
