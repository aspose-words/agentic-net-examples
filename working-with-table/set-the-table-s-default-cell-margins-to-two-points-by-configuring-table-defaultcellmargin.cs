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

        // Start a table.
        Table table = builder.StartTable();

        // Ensure the table has at least one row before applying table‑level padding.
        table.EnsureMinimum();

        // Set the default cell margins (padding) for the table – 2 points on each side.
        table.LeftPadding = 2;
        table.RightPadding = 2;
        table.TopPadding = 2;
        table.BottomPadding = 2;

        // Build a simple 2×2 table to demonstrate the margins.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Verify that the margins were applied correctly.
        if (table.LeftPadding != 2 ||
            table.RightPadding != 2 ||
            table.TopPadding != 2 ||
            table.BottomPadding != 2)
        {
            throw new InvalidOperationException("Default cell margins were not set correctly.");
        }

        // Save the document.
        doc.Save("DefaultCellMargin.docx");
    }
}
