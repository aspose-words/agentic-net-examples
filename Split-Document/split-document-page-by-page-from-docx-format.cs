using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("InputDocument.docx");

        // Determine the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Loop through each page and extract it as a separate document.
        for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            // Extract a single page (zero‑based index, count = 1).
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Build a filename for the extracted page.
            string outputPath = $"Page_{pageIndex + 1}.docx";

            // Save the extracted page as a new DOCX file.
            pageDoc.Save(outputPath);
        }
    }
}
