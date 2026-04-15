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

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Insert a caption paragraph directly below the table using the built‑in "Caption" style.
            // The caption will be numbered automatically via a SEQ field.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;

            // Create a bookmark that will be referenced later.
            const string bookmarkName = "Table1";
            builder.StartBookmark(bookmarkName);

            // Write the caption label and number.
            builder.Write("Table ");
            // Insert a SEQ field for automatic numbering.
            builder.InsertField("SEQ Table \\* ARABIC");
            builder.Write(" Sample table created with Aspose.Words.");

            // End the bookmark.
            builder.EndBookmark(bookmarkName);

            // Reset paragraph formatting for subsequent text.
            builder.ParagraphFormat.ClearFormatting();

            // Insert a paragraph that references the table caption.
            builder.Writeln();
            builder.Write("Refer to ");

            // Insert a cross‑reference to the bookmark (the caption number will be displayed).
            builder.InsertField($"REF {bookmarkName} \\h");
            builder.Writeln(" for more details.");

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
            doc.Save(outputPath);
        }
    }
}
