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

        // Build a simple 3x2 table.
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

        // At this point the table has 3 rows.
        // Delete the second row (index 1) using the Rows collection.
        table.Rows[1].Remove();

        // Optional validation: the table should now contain 2 rows.
        if (table.Rows.Count != 2)
            throw new InvalidOperationException("Row removal failed.");

        // Save the document to the local file system.
        string outputPath = "DeletedRowTable.docx";
        doc.Save(outputPath);
    }
}
