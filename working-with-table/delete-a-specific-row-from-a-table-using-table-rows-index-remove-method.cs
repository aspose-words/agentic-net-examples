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

        // Build a 3‑row, 2‑column table.
        Table table = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Write("Row 3, Cell 1");
        builder.InsertCell();
        builder.Write("Row 3, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Delete the second row (index 1) using the Rows collection.
        if (table.Rows.Count > 1)
        {
            table.Rows[1].Remove();
        }

        // Optional: display the remaining row count.
        Console.WriteLine($"Rows after deletion: {table.Rows.Count}");

        // Save the document to the local file system.
        doc.Save("DeletedRowTable.docx");
    }
}
