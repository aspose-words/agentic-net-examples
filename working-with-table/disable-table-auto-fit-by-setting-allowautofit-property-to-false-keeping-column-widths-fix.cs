using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAutoFit
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table.
            Table table = builder.StartTable();

            // First row, first cell with a fixed width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write("Fixed width column 1");

            // First row, second cell with a fixed width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Write("Fixed width column 2");

            // End the first row.
            builder.EndRow();

            // Second row, first cell (widths remain fixed).
            builder.InsertCell();
            builder.Write("Row 2, col 1");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Row 2, col 2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Disable automatic resizing of cells; keep column widths fixed.
            table.AllowAutoFit = false;

            // Save the document to a file.
            doc.Save("TableAllowAutoFit.docx");
        }
    }
}
