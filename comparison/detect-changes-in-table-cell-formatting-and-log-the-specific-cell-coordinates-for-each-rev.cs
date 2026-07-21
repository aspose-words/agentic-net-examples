using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the original document with a simple 2x2 table.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("A2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("B1");
        builder.InsertCell();
        builder.Write("B2");
        builder.EndTable();

        // Clone the original to create a revised version.
        Document revised = (Document)original.Clone(true);

        // Change the shading (background color) of the cell at row 1, column 0 (second row, first column).
        Table table = revised.FirstSection.Body.Tables[0];
        Row targetRow = table.Rows[1];
        Cell targetCell = targetRow.Cells[0];
        targetCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

        // Compare the documents – revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Iterate through revisions and log format changes that affect table cells.
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange &&
                rev.ParentNode != null &&
                rev.ParentNode.NodeType == NodeType.Cell)
            {
                // The parent node of the revision is a Cell.
                Cell changedCell = (Cell)rev.ParentNode;

                // Obtain the containing Row and Table via the ParentRow and ParentTable properties.
                Row row = changedCell.ParentRow;
                Table tbl = row.ParentTable;

                // Determine zero‑based row and column indices.
                int rowIndex = tbl.Rows.IndexOf(row);
                int columnIndex = row.Cells.IndexOf(changedCell);

                Console.WriteLine($"Format change detected at row {rowIndex}, column {columnIndex}.");
            }
        }

        // Save the compared document (contains the revisions).
        original.Save("Compared.docx");
    }
}
