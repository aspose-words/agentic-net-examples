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

        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");

        builder.InsertCell();
        builder.Writeln("Cell 4");

        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table (the last table in the document).
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[doc.GetChildNodes(NodeType.Table, true).Count - 1];

        // Set spacing before and after the table using the available padding properties.
        // These properties control the space above and below the table.
        table.TopPadding = 12f;    // Space before the table (12 points).
        table.BottomPadding = 12f; // Space after the table (12 points).

        // Define output file name.
        string outputPath = "TableSpacing.docx";

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
