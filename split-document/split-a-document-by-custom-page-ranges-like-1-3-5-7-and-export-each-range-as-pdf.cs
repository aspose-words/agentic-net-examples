using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample source document with at least 7 pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is the content of page {i}.");
            if (i < 7) // Add a page break after each page except the last.
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document for reference (optional).
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Define custom page ranges to extract.
        string[] pageRanges = { "1-3", "5-7" };

        // Ensure the document layout is up‑to‑date so PageCount is accurate.
        int totalPages = sourceDoc.PageCount;

        foreach (string range in pageRanges)
        {
            // Parse the start and end page numbers.
            string[] parts = range.Split('-');
            if (parts.Length != 2 ||
                !int.TryParse(parts[0], out int startPage) ||
                !int.TryParse(parts[1], out int endPage))
            {
                throw new ArgumentException($"Invalid page range format: '{range}'. Expected format 'start-end'.");
            }

            if (startPage < 1 || endPage < startPage || endPage > totalPages)
                throw new ArgumentOutOfRangeException($"Page range '{range}' is out of bounds. Document has {totalPages} pages.");

            int pageCount = endPage - startPage + 1;

            // Aspose.Words uses zero‑based page indexes for extraction.
            int startIndexZeroBased = startPage - 1;

            // Extract the specified pages into a new document.
            Document splitDoc = sourceDoc.ExtractPages(startIndexZeroBased, pageCount);

            // Save the split document as PDF.
            string pdfPath = Path.Combine(outputDir, $"Part_{range}.pdf");
            splitDoc.Save(pdfPath);
        }

        // Validate that the expected PDF files were created.
        foreach (string range in pageRanges)
        {
            string pdfPath = Path.Combine(outputDir, $"Part_{range}.pdf");
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"Expected split PDF not found: {pdfPath}");
        }

        Console.WriteLine("Document split completed successfully.");
    }
}
