using System;
using Aspose.Words;

class SplitDocumentExample
{
    static void Main()
    {
        // Path to the source document.
        string sourceFile = @"C:\Docs\SourceDocument.docx";

        // Load the source document (uses the provided Document(string) constructor).
        Document sourceDoc = new Document(sourceFile);

        // Determine the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Iterate through each page and extract it as a separate document.
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // Extract a single page (pageIndex) from the source document.
            // This uses the Document.ExtractPages method, which has no specific rule, so free‑form code is allowed.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Define the output file name for the extracted page.
            string outFile = $@"C:\Docs\Extracted_Page_{pageIndex}.docx";

            // Save the extracted page as a DOCX file (uses the provided Document.Save(string) method).
            pageDoc.Save(outFile);

            // Load the newly saved DOCX file to demonstrate that it can be opened again.
            Document loadedPage = new Document(outFile);

            // Optional: verify that the loaded document contains the expected page text.
            Console.WriteLine($"Page {pageIndex} extracted and reloaded. Text preview:");
            Console.WriteLine(loadedPage.GetText().Trim().Substring(0, Math.Min(100, loadedPage.GetText().Trim().Length)));
        }
    }
}
