using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the source multi‑page DOCX document.
        Document sourceDoc = new Document(@"C:\Input\MultiPageDocument.docx");

        // Determine the total number of pages in the source document.
        int totalPages = sourceDoc.PageCount;

        // Iterate through each page and extract it as a separate document.
        for (int i = 0; i < totalPages; i++)
        {
            // Extract a single page (zero‑based index) from the source document.
            Document singlePageDoc = sourceDoc.ExtractPages(i, 1);

            // Build the output file name, e.g. Page_1.docx, Page_2.docx, …
            string outputPath = Path.Combine(
                @"C:\Output",
                $"Page_{i + 1}.docx");

            // Save the extracted page as an individual DOCX file.
            singlePageDoc.Save(outputPath, SaveFormat.Docx);
        }
    }
}
