using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace SplitMergedCellExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x3 table.
            builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Writeln("Cell 1");
            builder.InsertCell();
            builder.Writeln("Cell 2");
            builder.InsertCell();
            builder.Writeln("Cell 3");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Writeln("Cell 4");
            builder.InsertCell();
            builder.Writeln("Cell 5");
            builder.InsertCell();
            builder.Writeln("Cell 6");
            builder.EndRow();

            builder.EndTable();

            // Retrieve the table we just created.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Row firstRow = table.Rows[0];

            // Merge the second and third cells of the first row horizontally.
            Cell firstMergedCell = firstRow.Cells[1]; // This will become the "first" cell of the merge.
            firstMergedCell.CellFormat.HorizontalMerge = CellMerge.First;

            Cell secondMergedCell = firstRow.Cells[2];
            secondMergedCell.CellFormat.HorizontalMerge = CellMerge.Previous;

            // At this point the second and third cells are merged.
            // To split them back into separate cells, reset the merge settings.
            firstMergedCell.CellFormat.HorizontalMerge = CellMerge.None;
            secondMergedCell.CellFormat.HorizontalMerge = CellMerge.None;

            // Save the document to verify the result.
            string outputPath = "SplitMergedCell.docx";
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Confirmation message.
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
    }
}
