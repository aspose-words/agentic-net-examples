using System;
using Aspose.Words;

class SplitDocumentByPages
{
    static void Main()
    {
        // Load the source DOCX document.
        string sourcePath = "InputDocument.docx";
        Document sourceDoc = new Document(sourcePath);

        // Determine the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Extract each page individually and save it as a separate DOCX file.
        for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            // Extract a single page (zero‑based index, count = 1).
            Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Save the extracted page. File name includes the page number (1‑based for readability).
            string outputPath = $"Page_{pageIndex + 1}.docx";
            singlePageDoc.Save(outputPath);
        }
    }
}
