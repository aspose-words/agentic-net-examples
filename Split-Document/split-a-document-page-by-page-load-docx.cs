using System;
using Aspose.Words;

namespace DocumentPageSplitter
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            // The Document(string) constructor automatically detects the file format.
            Document sourceDoc = new Document("InputDocument.docx");

            // Ensure the layout is up‑to‑date so that the page count is accurate.
            sourceDoc.UpdatePageLayout();

            // Get the total number of pages in the document.
            int totalPages = sourceDoc.PageCount;

            // Iterate through each page, extract it, and save it as a separate DOCX file.
            for (int pageIndex = 0; pageIndex < totalPages; pageIndex++)
            {
                // ExtractPages uses a zero‑based page index and extracts the specified count of pages.
                Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Save the extracted page. The Save(string) method determines the format from the extension.
                string outputFileName = $"Page_{pageIndex + 1}.docx";
                singlePageDoc.Save(outputFileName);
            }
        }
    }
}
