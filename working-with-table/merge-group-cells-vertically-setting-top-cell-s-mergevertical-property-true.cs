using System;
using Aspose.Words;
using Aspose.Words.Tables;

class VerticalMergeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with three rows and two columns.
        builder.StartTable();

        // ---------- Row 1 ----------
        // First cell – start of the vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First; // top cell of the merged group
        builder.Write("This cell spans three rows vertically.");

        // Second column cell – normal (not merged).
        builder.InsertCell();
        builder.Write("Row 1, Column 2");
        builder.EndRow();

        // ---------- Row 2 ----------
        // First column cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous; // continue vertical merge
        // No content needed for merged cells; they must be empty.
        // Second column cell – normal.
        builder.InsertCell();
        builder.Write("Row 2, Column 2");
        builder.EndRow();

        // ---------- Row 3 ----------
        // First column cell – merge with the cell above (still part of the same group).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        // Second column cell – normal.
        builder.InsertCell();
        builder.Write("Row 3, Column 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("VerticalMerge.docx");
    }
}
