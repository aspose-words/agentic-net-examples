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

        // Add some surrounding text before the table.
        builder.Writeln("This is some surrounding text before the table.");

        // Build a simple 2x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set the table's text wrapping mode to "Around" the surrounding text.
        // The TextWrapping property controls how text flows around the table.
        table.TextWrapping = TextWrapping.Around;

        // Add some surrounding text after the table.
        builder.Writeln("This is some surrounding text after the table.");

        // Save the document.
        string outputPath = "WrapAroundTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Indicate success (non‑interactive).
        Console.WriteLine("Document created successfully.");
    }
}
