using System;
using Aspose.Words;

namespace AsposeWordsSplitAndLoad
{
    class Program
    {
        static void Main()
        {
            // Path to the source document that will be split.
            string sourcePath = "SourceDocument.docx";

            // Load the source document (DOCX) using the Document(string) constructor.
            Document sourceDoc = new Document(sourcePath);

            // Determine the total number of pages in the source document.
            int totalPages = sourceDoc.PageCount;

            // Iterate through each page, extract it as a separate document,
            // save the part as a DOCX file, and then load the saved part back.
            for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
            {
                // Extract a single page (pageIndex) from the source document.
                // ExtractPages returns a new Document containing the specified range.
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

                // Define a filename for the extracted page part.
                string partFileName = $"ExtractedPage_{pageIndex}.docx";

                // Save the extracted page as a DOCX file using the Save(string) method.
                pageDoc.Save(partFileName);

                // Load the saved DOCX part back into a Document object.
                Document loadedPart = new Document(partFileName);

                // At this point, 'loadedPart' contains the same content as the extracted page.
                // Additional processing can be performed here if needed.
                Console.WriteLine($"Page {pageIndex} saved and reloaded from '{partFileName}'.");
            }
        }
    }
}
