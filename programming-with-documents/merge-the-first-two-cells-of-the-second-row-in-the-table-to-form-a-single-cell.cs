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

        // Build a table with two rows and three columns.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.InsertCell();
        builder.Writeln("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.InsertCell();
        builder.Writeln("R2C3");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Merge the first two cells of the second row.
        Row secondRow = table.Rows[1];
        // Set the first cell to start a horizontal merge.
        secondRow.Cells[0].CellFormat.HorizontalMerge = CellMerge.First;
        // Set the second cell to continue the merge.
        secondRow.Cells[1].CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the document.
        doc.Save("MergedTable.docx");
    }
}
