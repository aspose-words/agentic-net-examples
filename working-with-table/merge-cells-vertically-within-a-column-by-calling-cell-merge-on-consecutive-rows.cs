using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 4x2 table.
        builder.StartTable();

        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Merge the first three cells in the first column vertically.
        // Set the first cell as the start of a vertical merge range.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.CellFormat.VerticalMerge = CellMerge.First;

        // Mark the next two cells as part of the same vertical merge range.
        table.Rows[1].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;
        table.Rows[2].Cells[0].CellFormat.VerticalMerge = CellMerge.Previous;

        // Save the resulting document.
        doc.Save("MergedCells.docx");
    }
}
