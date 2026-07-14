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

        // Build a 3‑column, 2‑row table.
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
        builder.EndTable(); // Ends the table.

        // Delete the second column (index 1).
        // Aspose.Words does not expose a Columns collection; remove the cell at the
        // desired index from each row instead.
        int columnIndexToRemove = 1;
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > columnIndexToRemove)
        {
            foreach (Row row in table.Rows)
            {
                // Ensure the row still has enough cells before removal.
                if (row.Cells.Count > columnIndexToRemove)
                {
                    row.Cells.RemoveAt(columnIndexToRemove);
                }
            }
        }

        // Save the modified document.
        doc.Save("DeletedColumn.docx");
    }
}
