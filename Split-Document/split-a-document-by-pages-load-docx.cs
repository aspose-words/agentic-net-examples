using System;
using Aspose.Words;

class DocumentSplitter
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("InputDocument.docx");

        // Determine the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Iterate through each page, extract it, and save as a separate DOCX file.
        for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            // Extract a single page (zero‑based index, count = 1).
            Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Save the extracted page to a new file. File name includes the page number (1‑based).
            string outputFileName = $"Page_{pageIndex + 1}.docx";
            singlePageDoc.Save(outputFileName);
        }
    }
}
