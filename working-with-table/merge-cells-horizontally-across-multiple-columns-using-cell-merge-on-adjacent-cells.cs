using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableCellMergeExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2‑row, 3‑column table.
            builder.StartTable();

            // First row cells.
            builder.InsertCell();
            builder.Writeln("Merged Cell");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.InsertCell();
            builder.Writeln("Cell 3");
            builder.EndRow();

            // Second row cells.
            builder.InsertCell();
            builder.Writeln("Row2 Cell1");
            builder.InsertCell();
            builder.Writeln("Row2 Cell2");
            builder.InsertCell();
            builder.Writeln("Row2 Cell3");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Retrieve the created table.
            Table table = doc.FirstSection.Body.Tables[0];

            // Merge the first three cells of the first row horizontally.
            Cell firstCell = table.Rows[0].Cells[0];
            firstCell.CellFormat.HorizontalMerge = CellMerge.First;
            table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;
            table.Rows[0].Cells[2].CellFormat.HorizontalMerge = CellMerge.Previous;

            // Save the document.
            string outputPath = "MergedTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
