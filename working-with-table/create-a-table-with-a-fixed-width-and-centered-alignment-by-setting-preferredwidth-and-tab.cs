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
        builder.Writeln("Cell 1,1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 1,2");

        // End first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Cell 2,2");

        // End second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Set the table to have a fixed width of 400 points.
        table.PreferredWidth = PreferredWidth.FromPoints(400);

        // Center the table on the page.
        table.Alignment = TableAlignment.Center;

        // Save the document.
        string outputPath = "TableFixedWidth.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
