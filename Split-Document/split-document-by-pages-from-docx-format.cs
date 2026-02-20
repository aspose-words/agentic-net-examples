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

        // Options for extracting pages.
        PageExtractOptions extractOptions = new PageExtractOptions();
        // Keep the original page numbering and NUMPAGES fields (default behavior).
        extractOptions.UpdatePageStartingNumber = true;
        extractOptions.UnlinkPagesNumberFields = true;

        // Loop through each page and save it as a separate DOCX file.
        for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            // Extract a single page (zero‑based index) from the source document.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1, extractOptions);

            // Build the output file name, e.g., "Page_1.docx", "Page_2.docx", ...
            string outputFileName = $"Page_{pageIndex + 1}.docx";

            // Save the extracted page as a new DOCX document.
            pageDoc.Save(outputFileName, SaveFormat.Docx);
        }
    }
}
