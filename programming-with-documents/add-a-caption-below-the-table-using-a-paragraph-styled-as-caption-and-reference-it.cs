using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCaptionExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple 2x2 table.
            builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            builder.EndTable();

            // Insert a caption below the table using the built‑in "Caption" style.
            // The caption is also wrapped in a bookmark so that we can reference it later.
            const string captionBookmark = "TableCaption";
            builder.StartBookmark(captionBookmark);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
            builder.Writeln("Table 1. Sample Table");
            builder.EndBookmark(captionBookmark);
            // Reset style to default for following paragraphs.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

            // Insert a paragraph that will contain a cross‑reference to the table caption.
            builder.Writeln();
            builder.Write("Reference to the table: ");
            // REF field with \\h creates a hyperlink and \\p inserts the page number.
            builder.InsertField($"REF {captionBookmark} \\h \\p");
            builder.Writeln();

            // Save the document to disk.
            const string outputPath = "Output.docx";
            doc.Save(outputPath);
        }
    }
}
