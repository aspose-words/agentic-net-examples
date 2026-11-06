using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAutoFitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Disable automatic resizing and fix column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Set a fixed width for each cell (e.g., 100 points).
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
                }
            }

            // Add long text to the first cell to demonstrate that the column width does not change.
            Cell targetCell = table.Rows[0].Cells[0];
            builder.MoveTo(targetCell.FirstParagraph);
            builder.Write(
                "This is a very long piece of text that would normally cause the column to expand if AutoFit were enabled. " +
                "Because AutoFit is disabled, the text will wrap within the fixed column width.");

            // Save the document.
            string outputPath = "TableAutoFit.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
