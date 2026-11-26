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

        // Build a simple 3‑row, 2‑column table.
        builder.StartTable();

        // Row 0
        builder.InsertCell();
        builder.Write("R0C0");
        builder.InsertCell();
        builder.Write("R0C1");
        builder.EndRow();

        // Row 1 (the row we will delete)
        builder.InsertCell();
        builder.Write("R1C0");
        builder.InsertCell();
        builder.Write("R1C1");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("R2C0");
        builder.InsertCell();
        builder.Write("R2C1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Remember original row count.
        int originalRowCount = table.Rows.Count;

        // Delete the second row (index 1).
        table.Rows[1].Remove();

        // Verify that the row count decreased by one.
        if (table.Rows.Count != originalRowCount - 1)
            throw new InvalidOperationException("Row removal failed.");

        // Save the modified document.
        string outputPath = "DeletedRow.docx";
        doc.Save(outputPath);

        // Confirm that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output file was not saved.", outputPath);
    }
}
