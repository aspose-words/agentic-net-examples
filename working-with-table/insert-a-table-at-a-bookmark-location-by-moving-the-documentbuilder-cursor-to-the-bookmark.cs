using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class InsertTableAtBookmark
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some introductory text.
        builder.Writeln("Document with a bookmark below:");
        // Insert a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Bookmark location");
        builder.EndBookmark("MyBookmark");

        // Save the intermediate document (optional, just for reference).
        string intermediatePath = "Sample.docx";
        doc.Save(intermediatePath);

        // Move the builder cursor to the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Ensure we are in a paragraph before starting the table.
        builder.Writeln(); // Creates a new empty paragraph at the bookmark position.

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the final document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed successfully.
        Console.WriteLine("Table inserted at bookmark and document saved as " + outputPath);
    }
}
