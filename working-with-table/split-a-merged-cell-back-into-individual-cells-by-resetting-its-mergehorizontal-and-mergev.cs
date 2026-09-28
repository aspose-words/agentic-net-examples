using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class SplitMergedCellExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Reference cells for merging.
        Cell topLeft = table.Rows[0].Cells[0];
        Cell topRight = table.Rows[0].Cells[1];
        Cell bottomLeft = table.Rows[1].Cells[0];
        Cell bottomRight = table.Rows[1].Cells[1];

        // Merge top-left and top-right horizontally.
        topLeft.CellFormat.HorizontalMerge = CellMerge.First;
        topRight.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Merge top-left and bottom-left vertically.
        topLeft.CellFormat.VerticalMerge = CellMerge.First;
        bottomLeft.CellFormat.VerticalMerge = CellMerge.Previous;

        // At this point the top-left cell spans two columns and two rows.

        // ----- Split the merged cell back into individual cells -----
        // Reset horizontal and vertical merge properties for all cells that were part of the merge.
        topLeft.CellFormat.HorizontalMerge = CellMerge.None;
        topLeft.CellFormat.VerticalMerge = CellMerge.None;

        topRight.CellFormat.HorizontalMerge = CellMerge.None;
        topRight.CellFormat.VerticalMerge = CellMerge.None;

        bottomLeft.CellFormat.HorizontalMerge = CellMerge.None;
        bottomLeft.CellFormat.VerticalMerge = CellMerge.None;

        bottomRight.CellFormat.HorizontalMerge = CellMerge.None;
        bottomRight.CellFormat.VerticalMerge = CellMerge.None;

        // Validate that all cells are now unmerged.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                if (cell.CellFormat.HorizontalMerge != CellMerge.None ||
                    cell.CellFormat.VerticalMerge != CellMerge.None)
                {
                    throw new InvalidOperationException("Cell merge properties were not reset correctly.");
                }
            }
        }

        // Save the document.
        string outputPath = "SplitMergedCell.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
