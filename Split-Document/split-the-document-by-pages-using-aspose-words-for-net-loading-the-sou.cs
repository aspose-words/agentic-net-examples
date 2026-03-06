using System;
using Aspose.Words;

namespace AsposeWordsSplitByPages
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document sourceDoc = new Document("SourceDocument.docx");

            // Get the total number of pages in the document.
            int totalPages = sourceDoc.PageCount;

            // Iterate through each page and extract it as a separate document.
            for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
            {
                // Extract a single page (pageIndex is 1‑based).
                Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

                // Save the extracted page as a new DOCX file.
                string outputFileName = $"Page_{pageIndex}.docx";
                singlePageDoc.Save(outputFileName);
            }
        }
    }
}
