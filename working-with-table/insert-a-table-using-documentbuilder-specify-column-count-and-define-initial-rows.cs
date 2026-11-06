using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the builder is positioned at the end of the document.
        builder.MoveToDocumentEnd();

        // Start a table with three columns.
        builder.StartTable();

        // First row – three cells.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 3");
        builder.EndRow();

        // Second row – three cells.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create output file: {outputPath}");
        }
    }
}
