using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDynamicInsert
{
    class Program
    {
        static void Main()
        {
            // Load the markdown template. Aspose.Words can detect the format from the file extension.
            Document template = new Document("Template.md");

            // Prepare a list of source documents that will be inserted into the template.
            List<Document> sourceDocs = new List<Document>
            {
                new Document("Section1.docx"),
                new Document("Section2.docx"),
                new Document("Section3.docx")
            };

            // Use a DocumentBuilder to navigate the template and insert the source documents.
            DocumentBuilder builder = new DocumentBuilder(template);

            // Optionally, place a bookmark in the markdown template where the insertion should occur.
            // For this example we assume a bookmark named "InsertHere" exists.
            // If the bookmark does not exist, the builder will stay at the document start.
            if (template.Range.Bookmarks["InsertHere"] != null)
                builder.MoveToBookmark("InsertHere");
            else
                builder.MoveToDocumentEnd();

            // Insert each source document sequentially.
            foreach (Document src in sourceDocs)
            {
                // InsertDocument keeps the original formatting of the source document.
                builder.InsertDocument(src, ImportFormatMode.KeepSourceFormatting);
                // Add a page break between inserted sections for readability.
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Save the final document. The format is inferred from the file extension.
            template.Save("Result.docx");
        }
    }
}
