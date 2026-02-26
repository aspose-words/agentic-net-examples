using System;
using Aspose.Words;

namespace AsposeWordsSplitDocument
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            Document sourceDoc = new Document("InputDocument.docx");

            // Determine the total number of pages in the source document.
            int totalPages = sourceDoc.PageCount;

            // Iterate through each page and extract it as a separate document.
            for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
            {
                // Extract a single page (zero‑based index, count = 1).
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Save the extracted page to a new DOCX file.
                string outputFileName = $"Page_{pageIndex + 1}.docx";
                pageDoc.Save(outputFileName);
            }
        }
    }
}
