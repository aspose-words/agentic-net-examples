using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two adjacent cells.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row row = table.Rows[0];
        Cell firstCell = row.Cells[0];
        Cell secondCell = row.Cells[1];

        // Merge the two cells horizontally.
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }
    }
}
