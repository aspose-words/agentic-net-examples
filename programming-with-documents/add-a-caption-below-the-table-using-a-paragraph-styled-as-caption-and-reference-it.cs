using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCaptionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bookmark that will surround the table so we can reference it later.
            builder.StartBookmark("MyTable");

            // Build a simple 2x2 table.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            builder.EndTable();

            // End the bookmark after the table.
            builder.EndBookmark("MyTable");

            // Insert a caption paragraph directly below the table.
            // Use the built‑in "Caption" style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
            builder.Writeln("Table 1: Sample table.");

            // Move the cursor to the start of the document to add a reference to the table.
            builder.MoveToDocumentStart();
            builder.Writeln("See Table ");
            // Insert a cross‑reference field that points to the bookmark "MyTable".
            // The \\h switch makes the reference a hyperlink.
            builder.InsertField(" REF MyTable \\h ");
            builder.Writeln(" for details.");

            // Ensure the output directory exists.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithCaption.docx");
            doc.Save(outputPath);
        }
    }
}
