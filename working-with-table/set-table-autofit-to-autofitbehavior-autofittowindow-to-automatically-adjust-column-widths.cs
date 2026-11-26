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

        // Set the table to automatically fit to the window (page margins).
        table.AutoFit(AutoFitBehavior.AutoFitToWindow);

        // Save the document to a file.
        string outputPath = "TableAutoFit.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
