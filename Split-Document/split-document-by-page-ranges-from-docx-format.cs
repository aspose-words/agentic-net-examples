using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

public class DocumentSplitter
{
    /// <summary>
    /// Splits the source DOCX into separate documents based on the supplied page ranges.
    /// </summary>
    /// <param name="sourcePath">Full path to the source DOCX file.</param>
    /// <param name="outputFolder">Folder where the split documents will be saved.</param>
    /// <param name="ranges">List of tuples where Item1 is the zero‑based start page index and Item2 is the number of pages to extract.</param>
    public static void SplitByPageRanges(string sourcePath, string outputFolder, List<(int start, int count)> ranges)
    {
        // Load the source document (lifecycle rule: load)
        Document sourceDoc = new Document(sourcePath);

        // Ensure the output folder ends with a directory separator
        if (!outputFolder.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString()))
            outputFolder += System.IO.Path.DirectorySeparatorChar;

        // Iterate over each range and extract the corresponding pages
        for (int i = 0; i < ranges.Count; i++)
        {
            var (start, count) = ranges[i];

            // ExtractPages uses zero‑based page index (feature rule)
            Document part = sourceDoc.ExtractPages(start, count);

            // Build a file name like "Part_1.docx", "Part_2.docx", etc.
            string partPath = System.IO.Path.Combine(outputFolder, $"Part_{i + 1}.docx");

            // Save the extracted part (lifecycle rule: save)
            part.Save(partPath);
        }
    }

    // Example usage
    public static void Main()
    {
        string source = @"C:\Docs\SourceDocument.docx";
        string output = @"C:\Docs\SplitParts";

        // Define page ranges: pages 1‑2, pages 3‑4, and page 5 alone.
        // Note: page indices are zero‑based, so page 1 is index 0.
        var ranges = new List<(int start, int count)>
        {
            (0, 2), // pages 1‑2
            (2, 2), // pages 3‑4
            (4, 1)  // page 5
        };

        SplitByPageRanges(source, output, ranges);
    }
}
