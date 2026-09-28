using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class HeaderTableExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start building a table in the header.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Writeln("Header Cell 4");

        // End the second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "HeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
