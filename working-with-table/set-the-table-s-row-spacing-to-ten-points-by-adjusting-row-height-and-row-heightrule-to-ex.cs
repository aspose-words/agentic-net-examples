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

        // Build a simple 2‑row, 1‑column table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set each row's height to exactly 10 points.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.Height = 10;               // Height in points.
            row.RowFormat.HeightRule = HeightRule.Exactly;
        }

        // Save the document to disk.
        string outputPath = "TableRowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");
    }
}
