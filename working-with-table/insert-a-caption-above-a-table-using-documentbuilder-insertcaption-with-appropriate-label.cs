using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCaptionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document and associate a builder with it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a caption above the table.
            // The caption consists of a SEQ field that automatically numbers tables.
            // The field text will be rendered as "Table 1" (or the next number if more tables are added).
            builder.InsertField("SEQ Table \\* ARABIC", "Table ");
            builder.Writeln(" My Table"); // Add a description after the number.

            // Build a simple 2x2 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Save the document to the local file system.
            doc.Save("TableWithCaption.docx");
        }
    }
}
