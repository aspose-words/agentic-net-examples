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

        // Build a simple table with two rows.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Access the first row and set its height to exactly 20 points.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row firstRow = table.Rows[0];
        firstRow.RowFormat.Height = 20; // Height in points.
        firstRow.RowFormat.HeightRule = HeightRule.Exactly; // Exact height rule.

        // Save the document to disk.
        string outputPath = "RowHeightExact.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
