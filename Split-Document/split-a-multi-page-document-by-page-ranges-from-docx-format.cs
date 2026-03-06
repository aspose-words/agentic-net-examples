using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentSplitter
{
    static void Main()
    {
        // Path to the source multi‑page DOCX file.
        string sourcePath = @"C:\Docs\Source.docx";

        // Folder where the split documents will be written.
        string outputDir = @"C:\Docs\Split";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Define the page ranges to extract.
        // Each tuple contains a zero‑based start page index and the number of pages to take.
        var ranges = new (int start, int count)[]
        {
            (0, 2), // pages 1‑2
            (2, 2), // pages 3‑4
            (4, 1)  // page 5
        };

        // Load the original document (lifecycle rule: load).
        Document source = new Document(sourcePath);

        // Optional: configure how page numbers are handled in the extracted parts.
        PageExtractOptions options = new PageExtractOptions
        {
            // Keep the original page numbering (set to false to start at 1 for each part).
            UpdatePageStartingNumber = false,
            // Preserve NUMPAGES fields as fields (true replaces them with constant values).
            UnlinkPagesNumberFields = false
        };

        // Loop through each defined range, extract the pages and save the result.
        for (int i = 0; i < ranges.Length; i++)
        {
            var (start, count) = ranges[i];

            // Extract the specified range of pages (lifecycle rule: extract).
            Document part = source.ExtractPages(start, count, options);

            // Construct a file name for the part, e.g., Part_1.docx.
            string outPath = Path.Combine(outputDir, $"Part_{i + 1}.docx");

            // Save the extracted document (lifecycle rule: save).
            part.Save(outPath);
        }
    }
}
