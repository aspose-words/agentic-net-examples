using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace TableOfContentsForTables
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TOC field that will list only entries with the "Table" label.
            // Switches: \h – hyperlink, \z – hide page numbers in web layout, \c "Table" – use the "Table" label.
            builder.InsertTableOfContents("\\h \\z \\c \"Table\"");
            builder.InsertBreak(BreakType.PageBreak);

            // Add several tables each preceded by a caption that uses the SEQ field with the "Table" identifier.
            for (int i = 1; i <= 3; i++)
            {
                // Insert a caption paragraph.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Caption;
                builder.Write("Table ");
                // Insert the SEQ field for the table number.
                FieldSeq seq = (FieldSeq)builder.InsertField(FieldType.FieldSequence, true);
                seq.SequenceIdentifier = "Table";
                builder.Write($": Sample table {i}");
                builder.Writeln(); // End of caption paragraph.

                // Return to normal paragraph style for the table content.
                builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

                // Build a simple 1‑row, 2‑column table.
                Table table = builder.StartTable();

                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");
                builder.EndRow();

                builder.EndTable();
                builder.Writeln(); // Add a blank line after each table.
            }

            // Update all fields (including the TOC) so that the entries are populated.
            doc.UpdateFields();

            // Save the document to the local file system.
            doc.Save("TableOfContentsForTables.docx");
        }
    }
}
