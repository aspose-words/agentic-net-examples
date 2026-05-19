using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with multiple pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("This is the content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the content of page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is the content of page 3.");

        // Ensure the layout is up‑to‑date so PageCount is accurate.
        int pageCount = sourceDoc.PageCount;

        // Split the document page by page and save each part as a PDF.
        for (int pageIndex = 1; pageIndex <= pageCount; pageIndex++)
        {
            // Extract a single page. ExtractPages expects a zero‑based start index and a count.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex - 1, 1);

            // Build the output file name.
            string pdfPath = Path.Combine(outputDir, $"Part_{pageIndex}.pdf");

            // Save the extracted page as PDF.
            pageDoc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create split PDF for page {pageIndex}.");
        }

        // Indicate completion.
        Console.WriteLine($"Document split into {pageCount} PDF part(s) in folder: {outputDir}");
    }
}
