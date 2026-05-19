using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace TableTocExample
{
    public class Program
    {
        public static void Main()
        {
            // Output file path
            string outputPath = "TableWithToc.docx";

            // Create a new blank document and a builder for it
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bookmark that will surround the table
            builder.StartBookmark("TableBookmark");

            // Build a simple 2x2 table
            Table table = builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // Finish the table
            builder.EndTable();

            // End the bookmark
            builder.EndBookmark("TableBookmark");

            // Move cursor to the start of the document to insert the TOC field
            builder.MoveToDocumentStart();

            // Insert a TOC field and set it to use the bookmark we created
            FieldToc toc = (FieldToc)builder.InsertField(FieldType.FieldTOC, true);
            toc.BookmarkName = "TableBookmark";
            toc.InsertHyperlinks = true; // optional: make entries clickable

            // Update fields so the TOC reflects the current content
            doc.UpdateFields();

            // Save the document
            doc.Save(outputPath);

            // Verify that the file was created
            if (!File.Exists(outputPath))
                throw new Exception("The document was not saved correctly.");
        }
    }
}
