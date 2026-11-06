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

        // Start a new table.
        builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Writeln("Sample cell");

        // End the first row.
        builder.EndRow();

        // Retrieve the first row of the table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Row firstRow = table.Rows[0];

        // Set the row height to exactly 20 points.
        firstRow.RowFormat.Height = 20.0;               // Height in points.
        firstRow.RowFormat.HeightRule = HeightRule.Exactly;

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "RowHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
