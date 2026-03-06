using System;
using Aspose.Words;

class SplitByPages
{
    static void Main()
    {
        // Load the source DOCX document.
        Document sourceDoc = new Document("Source.docx");

        // Get the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Loop through each page, extract it, and save as a separate DOCX file.
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // Extract the current page (pageIndex) into a new Document.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Save the extracted page to a file named "Page_{pageIndex}.docx".
            pageDoc.Save($"Page_{pageIndex}.docx");
        }
    }
}
