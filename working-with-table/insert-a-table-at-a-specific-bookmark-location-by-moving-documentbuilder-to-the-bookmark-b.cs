using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content and a bookmark where the table will be inserted.
        builder.Writeln("This is some introductory text.");
        builder.StartBookmark("TableBookmark");
        builder.Writeln("Position for the table.");
        builder.EndBookmark("TableBookmark");

        // Move the builder to the bookmark location.
        builder.MoveToBookmark("TableBookmark");

        // Build a 2x2 table at the bookmark.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("R1C1");
        builder.InsertCell();
        builder.Writeln("R1C2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("R2C1");
        builder.InsertCell();
        builder.Writeln("R2C2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableAtBookmark.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Output file not found: {outputPath}");
        }
    }
}
