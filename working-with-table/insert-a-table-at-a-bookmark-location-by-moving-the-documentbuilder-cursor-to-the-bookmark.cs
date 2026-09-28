using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAtBookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some text and a bookmark where the table will be inserted.
            builder.Writeln("This is a paragraph before the table.");
            builder.StartBookmark("InsertTableHere");
            builder.Writeln("Bookmark location.");
            builder.EndBookmark("InsertTableHere");

            // Move the builder to the bookmark.
            builder.MoveToBookmark("InsertTableHere");

            // Insert a table at the bookmark location.
            builder.StartTable();

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

            // End the table.
            builder.EndTable();

            // Save the document.
            string outputPath = "Output.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // Indicate success.
            Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
        }
    }
}
