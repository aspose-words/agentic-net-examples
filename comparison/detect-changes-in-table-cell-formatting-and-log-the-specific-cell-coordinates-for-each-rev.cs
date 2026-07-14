using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class DetectTableCellFormattingChanges
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

        // Save the original document (optional, for inspection).
        string originalPath = "Original.docx";
        original.Save(originalPath);

        // Clone the original to create a revised version.
        Document revised = (Document)original.Clone(true);

        // Change the formatting of the cell at row 1, column 1 (zero‑based indices).
        Table table = revised.FirstSection.Body.Tables[0];
        Cell targetCell = table.Rows[1].Cells[1];
        targetCell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;

        // Save the revised document (optional, for inspection).
        string revisedPath = "Revised.docx";
        revised.Save(revisedPath);

        // Compare the documents to generate revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Iterate through revisions and log cell format changes with coordinates.
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode.NodeType == NodeType.Cell)
            {
                Cell changedCell = (Cell)rev.ParentNode;
                Row parentRow = changedCell.ParentRow;
                Table parentTable = parentRow?.ParentTable;

                int rowIndex = parentTable?.Rows.IndexOf(parentRow) ?? -1;
                int columnIndex = parentRow?.Cells.IndexOf(changedCell) ?? -1;

                Console.WriteLine($"Cell format changed at row {rowIndex}, column {columnIndex}.");
            }
        }

        // Save the document that now contains the revisions.
        string resultPath = "ComparisonResult.docx";
        original.Save(resultPath);
    }
}
