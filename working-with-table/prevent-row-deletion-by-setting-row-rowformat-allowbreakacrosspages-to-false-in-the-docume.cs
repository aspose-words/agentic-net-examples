using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace PreventRowBreakExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and add a few rows with two cells each.
            Table table = builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            // Prevent this row from breaking across pages.
            builder.RowFormat.AllowBreakAcrossPages = false;
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            // Prevent this row from breaking across pages.
            builder.RowFormat.AllowBreakAcrossPages = false;
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("Row 3, Cell 1");
            builder.InsertCell();
            builder.Write("Row 3, Cell 2");
            // Prevent this row from breaking across pages.
            builder.RowFormat.AllowBreakAcrossPages = false;
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Alternatively, ensure all rows have the property set (in case rows were added elsewhere).
            foreach (Row row in table.Rows)
            {
                row.RowFormat.AllowBreakAcrossPages = false;
            }

            // Prepare output directory.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "PreventRowBreak.docx");
            doc.Save(outputPath);
        }
    }
}
