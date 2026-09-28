using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with a bookmark named "MyBookmark".
        builder.StartBookmark("MyBookmark");
        builder.Writeln("This is the bookmarked location.");
        builder.EndBookmark("MyBookmark");

        // Move the builder cursor to the bookmark.
        builder.MoveToBookmark("MyBookmark");

        // Insert a rectangle shape at the bookmark location.
        // Width = 100 points, Height = 50 points.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
