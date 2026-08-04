using System;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Set the compatibility level to Word 2016.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Use DocumentBuilder to construct a complex table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the outer table.
        builder.StartTable();

        // -----------------------------------------------------------------
        // Row 1: A header cell that spans two columns (horizontal merge).
        // -----------------------------------------------------------------
        builder.InsertCell();
        builder.RowFormat.Height = 30;
        builder.RowFormat.HeightRule = HeightRule.Exactly;
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Header spanning two columns");
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous; // Merge with previous cell.
        builder.EndRow();

        // Reset merge for subsequent rows.
        builder.CellFormat.HorizontalMerge = CellMerge.None;

        // -----------------------------------------------------------------
        // Row 2: Simple two‑cell row.
        // -----------------------------------------------------------------
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // -----------------------------------------------------------------
        // Row 3: First cell contains a nested table, second cell regular text.
        // -----------------------------------------------------------------
        builder.InsertCell(); // First cell of the outer row.

        // Start a nested table inside the first cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Nested 1");
        builder.EndTable(); // End of nested table.

        // Continue with the second cell of the outer row.
        builder.InsertCell();
        builder.Write("Outer cell 2");
        builder.EndRow();

        // Finish the outer table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("ComplexTable.docx");
    }
}
