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

        // Insert a bookmark that will mark the position where the table will be placed.
        builder.StartBookmark("MyTableBookmark");
        builder.Writeln("Text before the table.");
        builder.EndBookmark("MyTableBookmark");

        // Move the builder's cursor to the start of the bookmark (inside it) so that the table is inserted there.
        // Parameters: bookmark name, isStart = true (move to start), isAfter = true (position after the start tag).
        builder.MoveToBookmark("MyTableBookmark", true, true);

        // Build the table at the current cursor position.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Add some text after the table to demonstrate normal flow.
        builder.Writeln();
        builder.Writeln("Text after the table.");

        // Save the document to a file in the current directory.
        string outputPath = "TableAtBookmark.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
