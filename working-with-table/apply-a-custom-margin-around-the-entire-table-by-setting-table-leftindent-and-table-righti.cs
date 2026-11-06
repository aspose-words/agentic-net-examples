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

        // Insert first row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Insert second row with two cells.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply custom left indent (margin) to the table.
        // The value is in points; 20 points ≈ 0.28 inches.
        table.LeftIndent = 20.0;

        // NOTE: The task mentions setting Table.RightIndent, but according to the
        // hard‑ban rules we must not use Table.RightIndent as it is prohibited.
        // Therefore only the left indent is applied.

        // Save the document to disk.
        string outputPath = "TableMargin.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // The program finishes here without waiting for user input.
    }
}
