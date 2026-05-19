using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace PreventRowBreakAcrossPages
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 2‑column table.
            Table table = builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("R3C1");
            builder.InsertCell();
            builder.Write("R3C2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Prevent each row from breaking across pages.
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = false;
            }

            // Define an output folder and file name.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
            Directory.CreateDirectory(outputDir);
            string outputPath = Path.Combine(outputDir, "Table_NoBreakAcrossPages.docx");

            // Save the document.
            doc.Save(outputPath);
        }
    }
}
