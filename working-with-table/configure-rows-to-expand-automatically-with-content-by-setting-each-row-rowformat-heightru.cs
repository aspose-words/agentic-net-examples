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

        // First row – short content.
        builder.InsertCell();
        builder.Writeln("Short text");
        builder.InsertCell();
        builder.Writeln("Another short text");
        builder.EndRow();

        // Second row – longer content to demonstrate automatic height expansion.
        builder.InsertCell();
        builder.Writeln("This is a longer piece of text that should cause the row height to increase automatically based on its content.");
        builder.InsertCell();
        builder.Writeln("Short");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Configure each row to expand automatically with its content.
        foreach (Row row in doc.GetChildNodes(NodeType.Row, true))
        {
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document to disk.
        string outputPath = "TableAutoHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }
    }
}
