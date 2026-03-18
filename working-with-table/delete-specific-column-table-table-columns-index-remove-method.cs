using System;
using Aspose.Words;
using Aspose.Words.Tables;

class DeleteTableColumnExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑row, 3‑column table.
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
        builder.EndTable();

        // ----- Delete the second column (zero‑based index = 1) -----
        // Aspose.Words versions prior to 23.5 do not expose Table.RemoveColumn or Table.Columns[].Remove().
        // The portable way is to iterate over each row and remove the cell at the required column index.
        int columnIndexToRemove = 1;
        foreach (Row row in table.Rows)
        {
            // Ensure the row actually has enough cells (defensive programming).
            if (row.Cells.Count > columnIndexToRemove)
                row.Cells.RemoveAt(columnIndexToRemove);
        }
        // ------------------------------------------------------------

        // Save the modified document.
        doc.Save("DeletedColumn.docx");
    }
}
