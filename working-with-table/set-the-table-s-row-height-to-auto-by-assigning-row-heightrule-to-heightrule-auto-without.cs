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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set each row's height rule to Auto (no explicit height).
        foreach (Row row in table.Rows)
        {
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document.
        string outputPath = "TableRowHeightAuto.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
