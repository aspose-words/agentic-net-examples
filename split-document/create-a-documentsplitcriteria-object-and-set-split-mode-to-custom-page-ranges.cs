using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document with several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the source document (optional, just to have a file on disk).
        const string sourceFile = "Source.docx";
        doc.Save(sourceFile);
        Console.WriteLine($"Source document saved as '{sourceFile}'.");

        // Ensure layout is performed so that PageCount is accurate.
        doc.UpdatePageLayout();

        // Define custom page ranges: 1-3 and 5-7.
        int[][] customRanges = new int[][]
        {
            new int[] { 1, 3 },
            new int[] { 5, 7 }
        };

        // Extract each range and save as a separate document.
        for (int i = 0; i < customRanges.Length; i++)
        {
            int startPage = customRanges[i][0];
            int endPage = customRanges[i][1];

            // Validate page numbers are within the source document's page count.
            if (startPage < 1 || endPage > doc.PageCount || startPage > endPage)
                throw new InvalidOperationException(
                    $"Invalid page range {startPage}-{endPage} for document with {doc.PageCount} pages.");

            // Aspose.Words uses zero‑based page index and a page count, not start/end.
            int startIndex = startPage - 1;                     // zero‑based start page
            int pageCount = endPage - startPage + 1;            // number of pages to extract

            Document splitDoc = doc.ExtractPages(startIndex, pageCount);
            string splitFileName = $"Split_{i + 1}_Pages_{startPage}_to_{endPage}.docx";
            splitDoc.Save(splitFileName);
            Console.WriteLine($"Saved split document: {splitFileName}");
        }

        // Simple validation that the expected split files exist.
        for (int i = 0; i < customRanges.Length; i++)
        {
            int startPage = customRanges[i][0];
            int endPage = customRanges[i][1];
            string expectedFile = $"Split_{i + 1}_Pages_{startPage}_to_{endPage}.docx";

            if (!File.Exists(expectedFile))
                throw new FileNotFoundException($"Expected split file was not created: {expectedFile}");
        }

        Console.WriteLine("All split documents were created successfully.");
    }
}
