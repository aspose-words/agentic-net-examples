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

        // Build a simple one‑row, one‑cell table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        // Get the first row of the table.
        Row row = table.Rows[0];

        // Set the row height to exactly 20 points.
        row.RowFormat.Height = 20.0;                     // Height in points.
        row.RowFormat.HeightRule = HeightRule.Exactly; // Enforce exact height.

        // Save the document to disk.
        string outputPath = "RowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
