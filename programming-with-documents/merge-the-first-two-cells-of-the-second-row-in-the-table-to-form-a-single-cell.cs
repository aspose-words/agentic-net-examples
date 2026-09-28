using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First row (header) with three cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Second row with three cells – we will merge the first two cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the table we just created (first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Get the second row (index 1) and merge its first two cells.
        Row secondRow = table.Rows[1];
        // Set the first cell to start a horizontal merge.
        secondRow.Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
        // Set the second cell to continue the merge.
        secondRow.Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        string outputPath = "MergedTable.docx";
        doc.Save(outputPath);
    }
}
