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
            builder.Writeln("Document with a bookmark where a table will be inserted.");

            // Create a bookmark named "InsertTableHere".
            builder.StartBookmark("InsertTableHere");
            // The bookmark can contain placeholder text; it will be replaced by the table.
            builder.Writeln("<<Table will be inserted here>>");
            builder.EndBookmark("InsertTableHere");

            // Move the builder's cursor to the start of the bookmark.
            builder.MoveToBookmark("InsertTableHere");

            // Insert a 2x2 table at the bookmark location.
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

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception($"Failed to create the output file: {outputPath}");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved successfully to '{outputPath}'.");
        }
    }
}
