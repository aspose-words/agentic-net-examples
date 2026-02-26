using System;
using Aspose.Words;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the source DOCX document.
        Document source = new Document("InputDocument.docx");

        // Determine how many pages the document contains.
        int totalPages = source.PageCount;

        // Iterate over each page, extract it, and save it as a separate file.
        for (int i = 0; i < totalPages; i++)
        {
            // Extract a single page (zero‑based index) from the source document.
            Document singlePage = source.ExtractPages(i, 1);

            // Save the extracted page. The file name includes the page number (1‑based for readability).
            string outputFileName = $"Page_{i + 1}.docx";
            singlePage.Save(outputFileName);
        }
    }
}
