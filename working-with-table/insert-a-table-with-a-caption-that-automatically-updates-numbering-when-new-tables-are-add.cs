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
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert first table caption with a SEQ field that will auto‑number tables.
            builder.Write("Table ");
            builder.InsertField(" SEQ Table \\* ARABIC ");
            builder.Writeln(": First table caption.");

            // Build the first table.
            Table table1 = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Insert second table caption. The SEQ field will continue numbering.
            builder.Writeln(); // Add a blank line between tables.
            builder.Write("Table ");
            builder.InsertField(" SEQ Table \\* ARABIC ");
            builder.Writeln(": Second table caption.");

            // Build the second table.
            Table table2 = builder.StartTable();
            builder.InsertCell();
            builder.Write("A");
            builder.InsertCell();
            builder.Write("B");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("C");
            builder.InsertCell();
            builder.Write("D");
            builder.EndRow();
            builder.EndTable();

            // Update all fields so that the SEQ fields reflect the correct table numbers.
            doc.UpdateFields();

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithCaptions.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved successfully.");
        }
    }
}
