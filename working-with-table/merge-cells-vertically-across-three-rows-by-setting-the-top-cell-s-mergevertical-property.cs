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

        // Build a table with three rows and two columns.
        builder.StartTable();

        // ----- Row 1 -----
        builder.InsertCell();               // Cell (0,0) – will be merged vertically.
        builder.Write("Merged Cell");
        builder.InsertCell();               // Cell (0,1)
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // ----- Row 2 -----
        builder.InsertCell();               // Cell (1,0) – part of vertical merge.
        builder.Write("");                  // Placeholder text.
        builder.InsertCell();               // Cell (1,1)
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // ----- Row 3 -----
        builder.InsertCell();               // Cell (2,0) – part of vertical merge.
        builder.Write("");                  // Placeholder text.
        builder.InsertCell();               // Cell (2,1)
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Merge the first column cells vertically across three rows.
        Cell topCell = table.Rows[0].Cells[0];
        topCell.CellFormat.VerticalMerge = CellMerge.First;

        Cell middleCell = table.Rows[1].Cells[0];
        middleCell.CellFormat.VerticalMerge = CellMerge.Previous;

        Cell bottomCell = table.Rows[2].Cells[0];
        bottomCell.CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the document.
        string outputPath = "VerticalMerge.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
