using System;
using System.Collections.Generic;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Path to the source multi‑page DOCX document.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the document using the constructor that accepts a file name.
        Document sourceDoc = new Document(sourcePath);

        // Define the page ranges to split the document into.
        // Each tuple contains the zero‑based start page index and the number of pages to extract.
        // Example: (0, 2) extracts pages 1‑2, (2, 3) extracts pages 3‑5, etc.
        List<(int startIndex, int pageCount)> ranges = new List<(int, int)>
        {
            (0, 2),   // Pages 1‑2
            (2, 3),   // Pages 3‑5
            (5, 1)    // Page 6
        };

        // Iterate over the defined ranges, extract each part, and save it as a separate file.
        for (int i = 0; i < ranges.Count; i++)
        {
            var (startIndex, pageCount) = ranges[i];

            // Extract the specified range of pages.
            // The ExtractPages method returns a new Document containing only those pages.
            Document partDoc = sourceDoc.ExtractPages(startIndex, pageCount);

            // Build the output file name (e.g., Part_1.docx, Part_2.docx, ...).
            string outputPath = $@"C:\Docs\Part_{i + 1}.docx";

            // Save the extracted part using the Save method that determines format from the extension.
            partDoc.Save(outputPath);
        }
    }
}
