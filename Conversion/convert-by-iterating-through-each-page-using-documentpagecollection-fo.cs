using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("input.docx");

        // Rebuild the layout so that page information is accurate.
        sourceDoc.UpdatePageLayout();

        // Total number of pages in the document.
        int pageCount = sourceDoc.PageCount;

        // Iterate through each page and extract it as a separate document.
        for (int pageIndex = 1; pageIndex <= pageCount; pageIndex++)
        {
            // Extract a single page (pages are 1‑based).
            Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Save the extracted page as its own DOCX file.
            singlePageDoc.Save($"Page_{pageIndex}.docx");
        }
    }
}
