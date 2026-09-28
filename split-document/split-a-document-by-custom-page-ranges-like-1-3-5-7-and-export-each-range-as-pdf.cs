using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample source document with ten pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"This is the content of page {i}.");
            if (i < 10)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document for reference (optional).
        string sourcePath = Path.Combine(outputDir, "source.docx");
        sourceDoc.Save(sourcePath);

        // Define custom page ranges to split.
        string[] pageRanges = { "1-3", "5-7" };

        // Get total page count of the source document.
        int totalPages = sourceDoc.PageCount;

        // Process each range.
        foreach (string range in pageRanges)
        {
            // Parse start and end page numbers.
            string[] parts = range.Split('-');
            if (parts.Length != 2 ||
                !int.TryParse(parts[0], out int startPage) ||
                !int.TryParse(parts[1], out int endPage))
            {
                throw new ArgumentException($"Invalid page range format: {range}");
            }

            // Validate page numbers.
            if (startPage < 1 || endPage > totalPages || startPage > endPage)
                throw new ArgumentOutOfRangeException($"Page range {range} is out of bounds. Document has {totalPages} pages.");

            // Extract the specified pages (ExtractPages uses 1‑based page numbers).
            int pageCount = endPage - startPage + 1;
            Document extracted = sourceDoc.ExtractPages(startPage, pageCount);

            // Save the extracted range as PDF.
            string outputPath = Path.Combine(outputDir, $"output_{range}.pdf");
            extracted.Save(outputPath, SaveFormat.Pdf);

            // Verify the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }

        // Program completed successfully.
    }
}
