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

        // Start building a table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1, Row 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2, Row 1");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1, Row 2");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2, Row 2");

        // End the second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set a fixed width for the table (e.g., 300 points) and center alignment.
        table.PreferredWidth = PreferredWidth.FromPoints(300);
        table.Alignment = TableAlignment.Center;

        // Save the document to a file.
        string outputPath = "TableOutput.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
