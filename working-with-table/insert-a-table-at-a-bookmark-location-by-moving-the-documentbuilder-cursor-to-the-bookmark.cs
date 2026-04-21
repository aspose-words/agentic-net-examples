using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableAtBookmark
{
    public class Program
    {
        public static void Main()
        {
            // Path for the output document.
            string outputPath = "TableAtBookmark.docx";

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("This paragraph appears before the bookmark.");

            // Create a bookmark named "InsertHere".
            builder.StartBookmark("InsertHere");
            builder.Writeln("Bookmark location.");
            builder.EndBookmark("InsertHere");

            // Move the builder cursor to the bookmark.
            builder.MoveToBookmark("InsertHere");

            // Build a 2x2 table at the bookmark location.
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

            // Finish the table.
            builder.EndTable();

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
