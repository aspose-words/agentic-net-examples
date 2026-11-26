using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    // Simple struct to hold a page range.
    private struct PageRange
    {
        public int Start;
        public int End;

        public PageRange(int start, int end)
        {
            Start = start;
            End = end;
        }

        public override string ToString() => $"From {Start} to {End}";
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a sample document with ten pages.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 10; i++)
        {
            builder.Writeln($"This is page {i}");
            if (i < 10)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // -----------------------------------------------------------------
        // 2. Define custom page ranges that we want to split out.
        // -----------------------------------------------------------------
        List<PageRange> customRanges = new List<PageRange>
        {
            new PageRange(1, 3),
            new PageRange(5, 7)
        };

        Console.WriteLine("Custom page ranges:");
        foreach (var range in customRanges)
            Console.WriteLine($"  {range}");

        // -----------------------------------------------------------------
        // 3. Extract each range as a separate document and save it.
        // -----------------------------------------------------------------
        string baseFileName = "SampleDocument";
        string outputFolder = Directory.GetCurrentDirectory(); // current folder

        foreach (var range in customRanges)
        {
            // Validate that the requested range is within the source document page count.
            int pageCount = doc.PageCount;
            if (range.Start < 1 || range.End > pageCount || range.Start > range.End)
                throw new ArgumentOutOfRangeException($"Invalid page range: {range.Start}-{range.End}");

            // Extract the pages.
            // Document.ExtractPages expects a zero‑based start index and a page count.
            int startIndex = range.Start - 1;                     // zero‑based start
            int count = range.End - range.Start + 1;              // number of pages to extract
            Document splitDoc = doc.ExtractPages(startIndex, count);

            // Build a deterministic file name.
            string outPath = Path.Combine(outputFolder,
                $"{baseFileName}_Pages_{range.Start}-{range.End}.docx");

            // Save the split document.
            splitDoc.Save(outPath);

            // Verify that the file was created.
            if (!File.Exists(outPath))
                throw new Exception($"Failed to save split document for pages {range.Start}-{range.End}.");

            Console.WriteLine($"Saved split document: {outPath}");
        }

        // -----------------------------------------------------------------
        // 4. Save the original sample document (optional verification).
        // -----------------------------------------------------------------
        string originalPath = Path.Combine(outputFolder, $"{baseFileName}.docx");
        doc.Save(originalPath);
        if (!File.Exists(originalPath))
            throw new Exception("Failed to save the original sample document.");

        Console.WriteLine($"Original document saved: {originalPath}");
    }
}
