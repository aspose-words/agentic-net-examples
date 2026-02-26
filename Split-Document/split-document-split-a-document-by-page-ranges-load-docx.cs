using System;
using System.IO;
using Aspose.Words;

class SplitDocumentByPageRanges
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Directory where the split documents will be saved.
        string outputFolder = @"C:\Docs\SplitParts";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Define page ranges to extract.
        // Each entry: [zero‑based start page index, number of pages to extract].
        // Example: first two pages, then next two pages.
        int[][] pageRanges = new int[][]
        {
            new int[] { 0, 2 }, // Pages 1‑2
            new int[] { 2, 2 }  // Pages 3‑4
            // Add more ranges as needed.
        };

        // Iterate over the defined ranges, extract each part, and save it.
        for (int i = 0; i < pageRanges.Length; i++)
        {
            int startIndex = pageRanges[i][0];
            int pageCount = pageRanges[i][1];

            // Extract the specified range of pages.
            Document partDoc = sourceDoc.ExtractPages(startIndex, pageCount);

            // Build the output file name.
            string partPath = Path.Combine(outputFolder, $"Part_{i + 1}.docx");

            // Save the extracted part.
            partDoc.Save(partPath);
        }
    }
}
