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

        // Start a table.
        builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Writeln("R3C1");
        builder.InsertCell();
        builder.Writeln("R3C2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Adjust spacing for each row.
        // Aspose.Words does not expose SpaceBefore/SpaceAfter on RowFormat.
        // Instead, we can control the row height to simulate spacing.
        // Height is measured in points; HeightRule.Exactly forces the row to that height.

        // Row 1: simulate larger spacing by setting a taller row height.
        Row row1 = table.Rows[0];
        row1.RowFormat.Height = 12.0;               // points
        row1.RowFormat.HeightRule = HeightRule.Exactly;

        // Row 2: medium spacing.
        Row row2 = table.Rows[1];
        row2.RowFormat.Height = 8.0;                // points
        row2.RowFormat.HeightRule = HeightRule.Exactly;

        // Row 3: smaller spacing.
        Row row3 = table.Rows[2];
        row3.RowFormat.Height = 6.0;                // points
        row3.RowFormat.HeightRule = HeightRule.Exactly;

        // Save the document.
        string outputPath = "RowSpacingExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
