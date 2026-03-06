using System;
using Aspose.Words;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load the primary document.
            Document mainDoc = new Document("input.docx");

            // Clone the loaded document (deep copy).
            Document clonedDoc = mainDoc.Clone();

            // Load a document that will be inserted at a specific position.
            Document docToInsert = new Document("insert.docx");

            // Use DocumentBuilder to move the cursor and insert the document.
            DocumentBuilder builder = new DocumentBuilder(mainDoc);
            builder.MoveToDocumentEnd();                     // Position at the end of the main document.
            builder.InsertBreak(BreakType.PageBreak);        // Optional break before insertion.
            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

            // Load a document that will be appended to the end of the main document.
            Document docToAppend = new Document("append.docx");
            mainDoc.AppendDocument(docToAppend, ImportFormatMode.UseDestinationStyles);

            // Save the combined document.
            mainDoc.Save("merged.docx");

            // -----------------------------------------------------------------
            // Split the merged document into separate files, one per page.
            // -----------------------------------------------------------------
            int totalPages = mainDoc.PageCount; // PageCount is calculated after layout.

            for (int page = 1; page <= totalPages; page++)
            {
                // Extract a single page as a new Document.
                Document pageDoc = mainDoc.ExtractPages(page, page);

                // Save the extracted page. File name includes the page number.
                string pageFileName = $"page_{page}.docx";
                pageDoc.Save(pageFileName);
            }

            // Optional: demonstrate that the cloned document is independent.
            // For example, add a paragraph to the clone and save it.
            DocumentBuilder cloneBuilder = new DocumentBuilder(clonedDoc);
            cloneBuilder.Writeln("This paragraph is added to the cloned document.");
            clonedDoc.Save("cloned.docx");
        }
    }
}
