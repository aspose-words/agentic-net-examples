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

        // Build a simple 2‑cell table.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        builder.Write("Cell 1");
        // End the first row.
        builder.EndRow();

        // Second cell (same row).
        builder.InsertCell();
        builder.Write("Cell 2");
        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Set the height of the first row to exactly 20 points.
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.Height = 20.0;                     // Height in points.
        firstRow.RowFormat.HeightRule = HeightRule.Exactly;   // Exact height rule.

        // Save the document.
        string outputPath = "RowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
