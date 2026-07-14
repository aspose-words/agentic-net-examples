using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row (the row we will delete).
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
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Delete the second row (index 1) using Table.Rows[index].Remove().
        if (table.Rows.Count > 1)
        {
            table.Rows[1].Remove();
        }

        // Simple validation: the table should now have 2 rows.
        if (table.Rows.Count != 2)
        {
            throw new InvalidOperationException("Row deletion failed; unexpected row count.");
        }

        // Save the document.
        string outputPath = "DeletedRowTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not saved.", outputPath);
        }

        // Indicate success (optional, not required for non‑interactive execution).
        Console.WriteLine("Table row deleted and document saved successfully.");
    }
}
