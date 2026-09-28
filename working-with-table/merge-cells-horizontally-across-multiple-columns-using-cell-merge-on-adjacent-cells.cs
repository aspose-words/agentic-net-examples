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

        // Build a simple 2‑row, 3‑column table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("A1");
        builder.InsertCell();
        builder.Writeln("B1");
        builder.InsertCell();
        builder.Writeln("C1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("A2");
        builder.InsertCell();
        builder.Writeln("B2");
        builder.InsertCell();
        builder.Writeln("C2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Merge the first two cells of the first row horizontally.
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[0].Cells[1];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Merge the last two cells of the second row horizontally.
        Cell thirdCell = table.Rows[1].Cells[1];
        Cell fourthCell = table.Rows[1].Cells[2];
        thirdCell.CellFormat.HorizontalMerge = CellMerge.First;
        fourthCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        string outputPath = "MergedCells.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
