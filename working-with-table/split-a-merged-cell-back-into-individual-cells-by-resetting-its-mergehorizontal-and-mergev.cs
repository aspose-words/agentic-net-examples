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

        // Build a 2x2 table.
        builder.StartTable();

        // ---- First row ----
        // Cell (0,0) will be merged horizontally and vertically.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;   // Start horizontal merge.
        builder.CellFormat.VerticalMerge = CellMerge.First;     // Start vertical merge.
        builder.Writeln("Merged Cell");

        // Cell (0,1) continues the horizontal merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Continue horizontal merge.
        // No text needed for the merged part.
        builder.EndRow();

        // ---- Second row ----
        // Cell (1,0) continues the vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;    // No horizontal merge here.
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // Continue vertical merge.
        // No text needed for the merged part.

        // Cell (1,1) is a normal cell.
        builder.InsertCell();
        builder.Writeln("Normal Cell");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Split the previously merged cells by resetting merge properties.
        // Reset horizontal and vertical merge for each cell in the table.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.HorizontalMerge = CellMerge.None;
                cell.CellFormat.VerticalMerge = CellMerge.None;
            }
        }

        // Save the resulting document.
        string outputPath = "SplitMergedCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
