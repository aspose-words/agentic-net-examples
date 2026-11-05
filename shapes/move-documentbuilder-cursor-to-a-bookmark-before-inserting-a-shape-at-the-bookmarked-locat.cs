using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Insert a paragraph with a bookmark where the shape will be placed.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartBookmark("MyBookmark");
        builder.Writeln("Text before the shape.");
        builder.EndBookmark("MyBookmark");

        // Move the builder cursor to the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Insert a rectangle shape at the bookmark location.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, clean up (not required for this example).
        // Console.WriteLine("Document created successfully.");
    }
}
