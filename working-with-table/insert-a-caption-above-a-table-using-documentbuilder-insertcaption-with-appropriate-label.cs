using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCaption
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a caption above the table.
            // The field "SEQ Table \\* ARABIC" generates an auto‑incrementing number for the label "Table".
            // The caption will look like: "Table 1: Sample table caption".
            builder.InsertField("SEQ Table \\* ARABIC", "1");
            builder.Write(" Table: Sample table caption");
            builder.Writeln(); // Move to the next paragraph (above the table).

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

            // Ensure the output directory exists.
            string outputDir = Directory.GetCurrentDirectory();
            Directory.CreateDirectory(outputDir);

            // Define output path.
            string outputPath = Path.Combine(outputDir, "TableWithCaption.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
