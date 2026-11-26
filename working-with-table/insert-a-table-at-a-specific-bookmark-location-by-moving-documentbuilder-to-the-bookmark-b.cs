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

        // Insert a paragraph that contains a bookmark named "InsertHere".
        builder.StartBookmark("InsertHere");
        builder.Writeln("This is the location of the bookmark.");
        builder.EndBookmark("InsertHere");

        // Move the builder to the bookmark position.
        builder.MoveToBookmark("InsertHere");

        // Build a simple 2x2 table at the bookmark location.
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableAtBookmark.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
