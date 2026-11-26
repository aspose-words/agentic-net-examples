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

        // Move the builder to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Build a simple 2x2 table inside the header.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "HeaderTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }

        // Optionally, indicate success (no interactive prompts required).
        Console.WriteLine("Document with header table created successfully.");
    }
}
