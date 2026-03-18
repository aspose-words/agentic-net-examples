using System;
using Aspose.Words;
using Aspose.Words.Tables;

class DeleteTableRowExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2‑row table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, cell 1.");
        builder.InsertCell();
        builder.Write("Row 1, cell 2.");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, cell 1.");
        builder.InsertCell();
        builder.Write("Row 2, cell 2.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Delete the first row (index 0) using the Row.Remove() method.
        int rowIndexToDelete = 0;
        if (rowIndexToDelete >= 0 && rowIndexToDelete < table.Rows.Count)
        {
            table.Rows[rowIndexToDelete].Remove();
        }

        // Save the resulting document.
        doc.Save("DeletedRow.docx");
    }
}
