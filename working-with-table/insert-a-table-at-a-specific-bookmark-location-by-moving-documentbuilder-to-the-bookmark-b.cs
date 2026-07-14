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
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("This paragraph appears before the bookmark.");

            // Insert a bookmark that will serve as the insertion point for the table.
            builder.StartBookmark("TableBookmark");
            builder.Writeln("Bookmark location – the table will be inserted here.");
            builder.EndBookmark("TableBookmark");

            // Add more text after the bookmark.
            builder.Writeln("This paragraph appears after the bookmark.");

            // Move the builder's cursor to the start of the bookmark.
            bool found = builder.MoveToBookmark("TableBookmark");
            if (!found)
                throw new InvalidOperationException("Bookmark 'TableBookmark' was not found.");

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
            builder.EndTable();

            // Save the document.
            string outputPath = "TableAtBookmark.docx";
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException($"Failed to create the output file: {outputPath}");
        }
    }
}
