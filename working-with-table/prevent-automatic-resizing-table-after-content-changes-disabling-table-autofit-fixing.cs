using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace TableAutoFitDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // ----- Define fixed column widths -----
            // First column – 120 points.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
            builder.Write("Header 1");

            // Second column – 80 points.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(80);
            builder.Write("Header 2");

            // End the header row.
            builder.EndRow();

            // ----- Add a data row with long text to test auto‑fit behavior -----
            builder.InsertCell();
            // The cell inherits the column width set above.
            builder.Write("This is a very long piece of text that would normally cause the column to expand if AutoFit were enabled.");

            builder.InsertCell();
            builder.Write("Short text");

            builder.EndRow();

            // End the table construction.
            builder.EndTable();

            // Disable automatic resizing of the table.
            table.AllowAutoFit = false;

            // Optionally, also call AutoFit with FixedColumnWidths to ensure the table layout is updated.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document.
            doc.Save("Table_NoAutoFit.docx", SaveFormat.Docx);
        }
    }
}
