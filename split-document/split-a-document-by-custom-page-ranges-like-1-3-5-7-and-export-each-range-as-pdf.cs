using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a sample document with 7 pages.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is the content of page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define custom page ranges to split.
        string rangeSpec = "1-3,5-7";
        List<(int Start, int End)> ranges = ParseRanges(rangeSpec);

        // Verify the source document has enough pages.
        int totalPages = sourceDoc.PageCount;
        if (totalPages == 0)
            throw new InvalidOperationException("Source document contains no pages.");

        // Process each range.
        int rangeIndex = 1;
        foreach (var (start, end) in ranges)
        {
            if (start < 1 || end > totalPages || start > end)
                throw new ArgumentOutOfRangeException($"Invalid page range: {start}-{end}");

            int pageCount = end - start + 1;
            // Aspose.Words uses zero‑based page indexes for extraction.
            int startIndexZeroBased = start - 1;
            Document extracted = sourceDoc.ExtractPages(startIndexZeroBased, pageCount);

            string outputFileName = $"output_{rangeIndex}_{start}-{end}.pdf";
            extracted.Save(outputFileName, SaveFormat.Pdf);

            // Validate that the file was created.
            if (!File.Exists(outputFileName))
                throw new IOException($"Failed to create output file: {outputFileName}");

            rangeIndex++;
        }
    }

    // Parses a string like "1-3,5-7" into a list of (start, end) tuples.
    private static List<(int Start, int End)> ParseRanges(string spec)
    {
        var result = new List<(int Start, int End)>();
        if (string.IsNullOrWhiteSpace(spec))
            return result;

        string[] parts = spec.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string part in parts)
        {
            string[] bounds = part.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (bounds.Length != 2)
                throw new FormatException($"Invalid range format: '{part}'");

            if (!int.TryParse(bounds[0].Trim(), out int start))
                throw new FormatException($"Invalid start page in range: '{bounds[0]}'");
            if (!int.TryParse(bounds[1].Trim(), out int end))
                throw new FormatException($"Invalid end page in range: '{bounds[1]}'");

            result.Add((start, end));
        }
        return result;
    }
}
