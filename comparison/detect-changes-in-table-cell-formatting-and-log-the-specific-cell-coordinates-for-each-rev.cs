using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableCellFormattingComparison
{
    public static void Main()
    {
        // Create the original document with a simple 2x2 table.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.StartTable();
        builderOriginal.InsertCell();
        builderOriginal.Write("Cell 1,1");
        builderOriginal.InsertCell();
        builderOriginal.Write("Cell 1,2");
        builderOriginal.EndRow();
        builderOriginal.InsertCell();
        builderOriginal.Write("Cell 2,1");
        builderOriginal.InsertCell();
        builderOriginal.Write("Cell 2,2");
        builderOriginal.EndTable();

        // Save the original document (optional, for inspection).
        original.Save("Original.docx");

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);
        DocumentBuilder builderRevised = new DocumentBuilder(revised);

        // Change the background shading of the cell at row 2, column 1.
        Table table = revised.FirstSection.Body.Tables[0];
        Row secondRow = table.Rows[1];
        Cell targetCell = secondRow.Cells[0];
        targetCell.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Yellow;

        // Save the revised document (optional, for inspection).
        revised.Save("Revised.docx");

        // Compare the original document with the revised one.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Iterate over revisions to find format changes in table cells.
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode is Cell cell)
            {
                // Determine the row index (1‑based).
                int rowIndex = table.Rows.IndexOf(cell.ParentRow) + 1;

                // Determine the column index (1‑based) within its row.
                int columnIndex = cell.ParentRow.Cells.IndexOf(cell) + 1;

                Console.WriteLine($"Format change detected in cell at Row {rowIndex}, Column {columnIndex}.");
            }
        }

        // Save the document that now contains the revisions.
        original.Save("ComparisonResult.docx");
    }
}
