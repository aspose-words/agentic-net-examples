using System;
using System.IO;
using Aspose.Words;

namespace TableAtBookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph with a bookmark named "MyBookmark".
            builder.Writeln("This is a paragraph before the bookmark.");
            builder.StartBookmark("MyBookmark");
            builder.Writeln("Bookmark location.");
            builder.EndBookmark("MyBookmark");
            builder.Writeln("This is a paragraph after the bookmark.");

            // Move the builder to the start of the bookmark.
            builder.MoveToBookmark("MyBookmark");

            // Insert a new paragraph after the bookmark to host the table.
            builder.Writeln();

            // Build a simple 2x2 table at the bookmark location.
            builder.StartTable();

            // First row
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();

            // Second row
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "TableAtBookmark.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
