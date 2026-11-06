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

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1, Row 1");
        builder.InsertCell();
        builder.Writeln("Cell 2, Row 1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 1, Row 2");
        builder.InsertCell();
        builder.Writeln("Cell 2, Row 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set each row to double line spacing by using Auto height rule and a specific height.
        foreach (Row row in table.Rows)
        {
            // HeightRule set to Auto.
            row.RowFormat.HeightRule = HeightRule.Auto;
            // Approximate double line spacing for a 12‑point font (24 points).
            row.RowFormat.Height = 24.0;
        }

        // Save the document to disk.
        string outputPath = "TableRowSpacing.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");
    }
}
