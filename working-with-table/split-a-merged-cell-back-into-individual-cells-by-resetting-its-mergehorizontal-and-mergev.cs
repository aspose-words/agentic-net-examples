using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table.
        builder.StartTable();

        // First row, first cell (will be merged horizontally and vertically).
        builder.InsertCell();
        builder.Writeln("Merged Cell");

        // First row, second cell (horizontal merge continuation).
        builder.InsertCell();
        // Get the cell that was just created and set its horizontal merge to Previous.
        Cell cellH = (Cell)builder.CurrentParagraph.ParentNode;
        cellH.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.Writeln(" ");

        // End first row.
        builder.EndRow();

        // Second row, first cell (vertical merge continuation).
        builder.InsertCell();
        // Get the cell that was just created and set its vertical merge to Previous.
        Cell cellV = (Cell)builder.CurrentParagraph.ParentNode;
        cellV.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Writeln(" ");

        // Second row, second cell (normal cell).
        builder.InsertCell();
        builder.Writeln("Cell 2,2");

        // End second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Locate the originally merged cell (first row, first cell).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell mergedCell = table.Rows[0].Cells[0];

        // Reset merge properties to split the cell back into individual cells.
        mergedCell.CellFormat.HorizontalMerge = CellMerge.None;
        mergedCell.CellFormat.VerticalMerge = CellMerge.None;

        // Also reset the continuation cells to ensure they are independent.
        table.Rows[0].Cells[1].CellFormat.HorizontalMerge = CellMerge.None;
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.None;

        // Save the document.
        string outputPath = "SplitMergedCell.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
