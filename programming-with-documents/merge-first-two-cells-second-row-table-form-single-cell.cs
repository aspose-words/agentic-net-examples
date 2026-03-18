using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table (3 rows × 3 columns) to demonstrate the merge.
        builder.StartTable();

        // ----- First row -----
        builder.InsertCell(); builder.Write("R1C1");
        builder.InsertCell(); builder.Write("R1C2");
        builder.InsertCell(); builder.Write("R1C3");
        builder.EndRow();

        // ----- Second row (the row we will merge) -----
        builder.InsertCell(); builder.Write("R2C1");
        builder.InsertCell(); builder.Write("R2C2");
        builder.InsertCell(); builder.Write("R2C3");
        builder.EndRow();

        // ----- Third row -----
        builder.InsertCell(); builder.Write("R3C1");
        builder.InsertCell(); builder.Write("R3C2");
        builder.InsertCell(); builder.Write("R3C3");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Access the second row (index 1 because rows are zero‑based).
        Row secondRow = table.Rows[1];

        // Merge the first two cells of the second row.
        // Mark the first cell as the start of a merged range.
        Cell firstCell = secondRow.Cells[0];
        firstCell.CellFormat.HorizontalMerge = CellMerge.First;

        // Mark the second cell to merge with the previous cell.
        Cell secondCell = secondRow.Cells[1];
        secondCell.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Save the resulting document.
        doc.Save("MergedCells.docx");
    }
}
