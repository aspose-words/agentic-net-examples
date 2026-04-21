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

        // Build a table with three rows and two columns.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row (the row we will delete later).
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndTable();

        // Verify initial row count (should be 3).
        if (table.Rows.Count != 3)
            throw new InvalidOperationException("Table should contain three rows before deletion.");

        // Delete the second row (index 1) using the Rows collection.
        Row rowToRemove = table.Rows[1];
        rowToRemove.Remove();

        // Verify that the row count is now 2.
        if (table.Rows.Count != 2)
            throw new InvalidOperationException("Row removal failed; table should contain two rows.");

        // Save the modified document.
        const string outputPath = "DeletedRowTable.docx";
        doc.Save(outputPath);
    }
}
