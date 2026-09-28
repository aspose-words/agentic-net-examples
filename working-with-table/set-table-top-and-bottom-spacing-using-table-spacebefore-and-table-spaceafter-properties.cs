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

        // Add some text before the table.
        builder.Writeln("Paragraph before the table.");

        // Set spacing before and after the table (in points) via the paragraph format.
        // The spacing is applied to the paragraph that will contain the table.
        builder.ParagraphFormat.SpaceBefore = 12; // 12 points before the table.
        builder.ParagraphFormat.SpaceAfter = 12;  // 12 points after the table.

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table (optional, shown for completeness).
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Save the document to a file.
        string outputPath = "TableSpacing.docx";
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
