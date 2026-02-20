using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("Input.docx");

        // Define the page ranges you want to split the document into.
        // Each tuple contains the start page (1‑based) and the end page (inclusive).
        (int start, int end)[] pageRanges = new (int, int)[]
        {
            (1, 2),   // First part: pages 1‑2
            (3, 4),   // Second part: pages 3‑4
            // Add more ranges as needed.
        };

        // Optional: configure how page numbers and NUMPAGES fields are handled
        // in the extracted documents.
        PageExtractOptions extractOptions = new PageExtractOptions
        {
            UpdatePageStartingNumber = true,   // Reset page numbering for each part.
            UnlinkPagesNumberFields = true     // Replace NUMPAGES fields with actual values.
        };

        // Extract each range and save it as a separate DOCX file.
        for (int i = 0; i < pageRanges.Length; i++)
        {
            var (startPage, endPage) = pageRanges[i];

            // Extract the specified page range.
            Document partDoc = sourceDoc.ExtractPages(startPage, endPage, extractOptions);

            // Save the extracted part.
            string outputPath = $"Output_Part_{i + 1}.docx";
            partDoc.Save(outputPath);
        }
    }
}
