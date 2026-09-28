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

        // Build a simple table with one row and one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the first table and its first row.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row row = table.Rows[0];

        // Set the row height rule to AtLeast and specify a minimum height of 15 points.
        row.RowFormat.HeightRule = HeightRule.AtLeast; // Use HeightRule enum.
        row.RowFormat.Height = 15; // Height is measured in points.

        // Save the document to a file.
        string outputPath = "TableRowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
