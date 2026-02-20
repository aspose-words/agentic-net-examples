using System;
using Aspose.Words;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the multi‑page DOCX document.
        string inputFile = @"C:\Docs\MultiPageDocument.docx";
        Document doc = new Document(inputFile);

        // Determine the total number of pages in the document.
        int pageCount = doc.PageCount;

        // Iterate through each page and extract it to a separate document.
        for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
        {
            // Extract a single page (zero‑based index) from the source document.
            Document pageDoc = doc.ExtractPages(pageIndex, 1);

            // Optional: customize page extraction behavior.
            // PageExtractOptions options = new PageExtractOptions();
            // options.UpdatePageStartingNumber = false;
            // options.UnlinkPagesNumberFields = false;
            // pageDoc = doc.ExtractPages(pageIndex, 1, options);

            // Save the extracted page as an individual DOCX file.
            string outputFile = $@"C:\Docs\Page_{pageIndex + 1}.docx";
            pageDoc.Save(outputFile);
        }
    }
}
