using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Folder for generated files
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with 7 pages
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document (optional, just for reference)
        string sourcePath = Path.Combine(outputDir, "Source.docx");
        sourceDoc.Save(sourcePath);

        // Define custom page ranges (1‑3 and 5‑7)
        var ranges = new List<(int start, int count)>
        {
            (1, 3), // pages 1 to 3
            (5, 3)  // pages 5 to 7
        };

        // Validate that the source document has enough pages
        if (sourceDoc.PageCount < 7)
            throw new InvalidOperationException("Source document does not contain the required number of pages.");

        // Process each range
        int partIndex = 1;
        foreach (var (startPage, pageCount) in ranges)
        {
            // Convert to zero‑based index for ExtractPages
            int zeroBasedStart = startPage - 1;

            // Extract the required pages
            Document part = sourceDoc.ExtractPages(zeroBasedStart, pageCount);

            // Prepare PDF save options (default options are sufficient)
            string pdfPath = Path.Combine(outputDir, $"Part_{partIndex}.pdf");
            part.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the file was created
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException($"Failed to create PDF for range {startPage}-{startPage + pageCount - 1}.", pdfPath);

            partIndex++;
        }

        // Final validation: ensure exactly two PDF files were created
        string[] pdfFiles = Directory.GetFiles(outputDir, "Part_*.pdf");
        if (pdfFiles.Length != ranges.Count)
            throw new InvalidOperationException("The expected number of split PDF files was not produced.");
    }
}
