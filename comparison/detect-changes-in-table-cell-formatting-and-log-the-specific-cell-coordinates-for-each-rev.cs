using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the original document with a 2x2 table.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.StartTable();

        // Row 1
        builderOriginal.InsertCell();
        builderOriginal.Writeln("Cell 1,1");
        builderOriginal.InsertCell();
        builderOriginal.Writeln("Cell 1,2");
        builderOriginal.EndRow();

        // Row 2
        builderOriginal.InsertCell();
        builderOriginal.Writeln("Cell 2,1");
        builderOriginal.InsertCell();
        builderOriginal.Writeln("Cell 2,2");
        builderOriginal.EndRow();

        builderOriginal.EndTable();

        // Apply shading to the top‑left cell in the original document.
        Table tableOriginal = original.FirstSection.Body.Tables[0];
        Cell originalCell = tableOriginal.Rows[0].Cells[0];
        originalCell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Save the original document (optional, for inspection).
        original.Save("Original.docx");

        // Create the revised document with the same table structure.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.StartTable();

        // Row 1
        builderRevised.InsertCell();
        builderRevised.Writeln("Cell 1,1");
        builderRevised.InsertCell();
        builderRevised.Writeln("Cell 1,2");
        builderRevised.EndRow();

        // Row 2
        builderRevised.InsertCell();
        builderRevised.Writeln("Cell 2,1");
        builderRevised.InsertCell();
        builderRevised.Writeln("Cell 2,2");
        builderRevised.EndRow();

        builderRevised.EndTable();

        // Change shading of the top‑left cell and add a border to the bottom‑right cell.
        Table tableRevised = revised.FirstSection.Body.Tables[0];
        Cell revisedCell00 = tableRevised.Rows[0].Cells[0];
        revisedCell00.CellFormat.Shading.BackgroundPatternColor = Color.Yellow; // Different shading.

        Cell revisedCell11 = tableRevised.Rows[1].Cells[1];
        revisedCell11.CellFormat.Borders.LineWidth = 2.0; // Add a visible border.

        // Save the revised document (optional, for inspection).
        revised.Save("Revised.docx");

        // Compare the documents. Revisions will be stored in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Iterate through revisions and log formatting changes in table cells.
        foreach (Revision revision in original.Revisions)
        {
            // Revision.ParentNode holds the node that was changed.
            if (revision.RevisionType == RevisionType.FormatChange && revision.ParentNode is Cell changedCell)
            {
                // Determine row and column indices (0‑based).
                Row parentRow = changedCell.ParentRow;
                Table parentTable = parentRow?.ParentTable;
                int rowIndex = parentTable?.Rows.IndexOf(parentRow) ?? -1;
                int columnIndex = parentRow?.Cells.IndexOf(changedCell) ?? -1;

                // Log the cell coordinates (1‑based for readability).
                Console.WriteLine($"Formatting revision detected at Row {rowIndex + 1}, Column {columnIndex + 1}.");
            }
        }

        // Save the document that contains the revisions.
        original.Save("TableFormattingComparison.docx");
    }
}
