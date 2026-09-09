using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the original document with a simple 2x3 table.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row}C{col}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Clone the original to create a revised version and change cell formatting.
        Document revised = (Document)original.Clone(true);
        Table table = (Table)revised.GetChild(NodeType.Table, 0, true);

        // Change formatting of two cells to generate format-change revisions.
        Cell cell01 = table.Rows[0].Cells[1]; // Row 0, Column 1
        cell01.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Yellow;

        Cell cell12 = table.Rows[1].Cells[2]; // Row 1, Column 2
        cell12.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightBlue;

        // Compare the documents. The original will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Log coordinates of each format-change revision that affects a table cell.
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode?.NodeType == NodeType.Cell)
            {
                Cell changedCell = (Cell)rev.ParentNode;
                Row parentRow = changedCell.ParentRow;
                Table parentTable = parentRow.ParentTable;

                int rowIndex = parentTable.Rows.IndexOf(parentRow);
                int columnIndex = parentRow.Cells.IndexOf(changedCell);

                Console.WriteLine($"Format change detected at row {rowIndex}, column {columnIndex}.");
            }
        }

        // Save the compared document for inspection.
        original.Save("ComparisonResult.docx");
    }
}
