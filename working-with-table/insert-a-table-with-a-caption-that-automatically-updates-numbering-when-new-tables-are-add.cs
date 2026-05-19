using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCaptionExample
{
    public class Program
    {
        public static void Main()
        {
            // Define output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithCaption.docx");

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

            // Insert a paragraph containing a SEQ field for automatic table numbering.
            // The field will produce "Table 1", "Table 2", etc., when the document is opened in Word.
            builder.Writeln(); // Ensure we are on a new paragraph.
            builder.InsertField("SEQ Table \\* ARABIC");
            builder.Write(" – Sample table caption.");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }
    }
}
