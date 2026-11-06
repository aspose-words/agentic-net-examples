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

        // Build a table with a horizontally merged cell in the first row.
        builder.StartTable();

        // First cell – start of merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // Second cell – continuation of merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.Write(string.Empty);

        builder.EndRow();

        // Second row – normal cells.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Locate the merged cells (first row, first two cells).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell firstCell = table.Rows[0].Cells[0];
        Cell secondCell = table.Rows[0].Cells[1];

        // Split the merged cell back into two separate cells by clearing the merge flags.
        firstCell.CellFormat.HorizontalMerge = CellMerge.None;
        secondCell.CellFormat.HorizontalMerge = CellMerge.None;

        // Save the document.
        string outputPath = "SplitCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
