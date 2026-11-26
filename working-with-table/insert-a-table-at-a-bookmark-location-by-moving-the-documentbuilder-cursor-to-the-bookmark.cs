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

        // Insert some initial text and a bookmark where the table will be placed.
        builder.Writeln("Document before the table.");
        builder.StartBookmark("InsertTableHere");
        builder.Writeln("Bookmark location.");
        builder.EndBookmark("InsertTableHere");

        // Move the builder cursor to the bookmark.
        builder.MoveToBookmark("InsertTableHere");

        // Build a simple 2x2 table at the bookmark location.
        builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        builder.EndTable();

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
