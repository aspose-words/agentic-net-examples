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

        // Insert a caption for the table (e.g., "Table 1") above the table.
        // Since DocumentBuilder.InsertCaption is not available, use a field to generate the number.
        builder.Write("Table ");
        builder.InsertField("SEQ Table \\* ARABIC");
        builder.Writeln(); // Move to the next line after the caption.

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row (header)
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Writeln("Cell A1");
        builder.InsertCell();
        builder.Writeln("Cell A2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "TableWithCaption.docx";
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
