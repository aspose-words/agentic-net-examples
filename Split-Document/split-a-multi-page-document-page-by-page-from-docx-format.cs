using System;
using Aspose.Words;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the multi‑page DOCX document.
        Document sourceDoc = new Document("InputDocument.docx");

        // Iterate through each page in the source document.
        for (int pageIndex = 0; pageIndex < sourceDoc.PageCount; pageIndex++)
        {
            // Extract a single page (zero‑based index, count = 1).
            Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Build an output file name like "Page_1.docx", "Page_2.docx", etc.
            string outputPath = $"Page_{pageIndex + 1}.docx";

            // Save the extracted page as a separate DOCX file.
            singlePageDoc.Save(outputPath);
        }
    }
}
