using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByRanges
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("Input.docx");

        // Define the page ranges to split the document into.
        // Each tuple contains the zero‑based start page index and the number of pages to extract.
        var pageRanges = new (int startIndex, int pageCount)[]
        {
            (0, 2), // Pages 1‑2
            (2, 3), // Pages 3‑5
            (5, 1)  // Page 6
        };

        // Iterate over the defined ranges, extract each part, and save it as a separate DOCX file.
        for (int i = 0; i < pageRanges.Length; i++)
        {
            var (start, count) = pageRanges[i];

            // Extract the specified pages into a new Document instance.
            Document extractedPart = sourceDoc.ExtractPages(start, count);

            // Optionally, customize page numbering or field handling using PageExtractOptions.
            // Example (uncomment if needed):
            // var extractOptions = new PageExtractOptions { UpdatePageStartingNumber = true, UnlinkPagesNumberFields = true };
            // extractedPart = sourceDoc.ExtractPages(start, count, extractOptions);

            // Save the extracted part to a new DOCX file.
            string outputPath = $"Output_Part{i + 1}.docx";
            extractedPart.Save(outputPath);
        }
    }
}
