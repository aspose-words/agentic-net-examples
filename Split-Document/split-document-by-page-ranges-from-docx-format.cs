using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByPageRanges
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("Input.docx");

        // Define the page ranges you want to split the document into.
        // Each inner array contains two integers: start page (1‑based) and end page (1‑based).
        int[][] pageRanges = new int[][]
        {
            new int[] { 1, 2 },   // Pages 1‑2
            new int[] { 3, 4 },   // Pages 3‑4
            // Add more ranges as needed.
        };

        // Options that control how the extracted pages are handled.
        PageExtractOptions extractOptions = new PageExtractOptions
        {
            // Keep the original page numbering in the extracted part.
            UpdatePageStartingNumber = true,
            // Replace NUMPAGES fields with their actual values.
            UnlinkPagesNumberFields = true
        };

        // Iterate over each defined range, extract the pages, and save them as separate documents.
        for (int i = 0; i < pageRanges.Length; i++)
        {
            int startPage = pageRanges[i][0]; // 1‑based page number.
            int endPage   = pageRanges[i][1]; // 1‑based page number.

            // ExtractPages uses zero‑based indices, so subtract 1 from each.
            Document partDoc = sourceDoc.ExtractPages(startPage - 1, endPage - 1, extractOptions);

            // Save the extracted part to a new DOCX file.
            string outputPath = $"Output_Part_{i + 1}.docx";
            partDoc.Save(outputPath);
        }
    }
}
