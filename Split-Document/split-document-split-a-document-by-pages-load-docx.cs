using System;
using System.IO;
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

        // Loop through each page and extract it to a separate document.
        for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
        {
            // Extract a single page. The ExtractPages method uses zero‑based page indices.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Optional: customize page numbering for the extracted document.
            // By default the start page number is preserved; uncomment to reset.
            // PageExtractOptions options = new PageExtractOptions
            // {
            //     UpdatePageStartingNumber = false,
            //     UnlinkPagesNumberFields = false
            // };
            // pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex, options);

            // Save the extracted page as a separate DOCX file.
            string outputPath = $"Page_{pageIndex + 1}.docx";
            pageDoc.Save(outputPath);
        }

        Console.WriteLine("Document split into individual pages successfully.");
    }
}
