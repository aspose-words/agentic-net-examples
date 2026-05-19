using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class DeleteTableColumnExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("R3C1");
        builder.InsertCell();
        builder.Write("R3C2");
        builder.InsertCell();
        builder.Write("R3C3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Verify initial column count.
        int initialColumnCount = table.Rows[0].Cells.Count;
        if (initialColumnCount != 3)
            throw new InvalidOperationException("Table should have 3 columns initially.");

        // Delete the second column (index 1) by removing the cell at that index from each row.
        int columnIndexToRemove = 1;
        foreach (Row row in table.Rows)
        {
            // Ensure the row has enough cells before removal.
            if (row.Cells.Count > columnIndexToRemove)
                row.Cells.RemoveAt(columnIndexToRemove);
        }

        // Verify column count after removal.
        int afterRemovalCount = table.Rows[0].Cells.Count;
        if (afterRemovalCount != initialColumnCount - 1)
            throw new InvalidOperationException("Column removal failed.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DeletedColumnTable.docx");
        doc.Save(outputPath);
    }
}
