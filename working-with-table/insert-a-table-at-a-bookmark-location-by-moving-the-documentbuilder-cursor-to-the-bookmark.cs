using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("This is a sample document.");

        // Create a bookmark that will serve as the insertion point for the table.
        builder.StartBookmark("InsertTableHere");
        builder.Writeln("Bookmark location.");
        builder.EndBookmark("InsertTableHere");

        // Move the builder's cursor to the start of the bookmark.
        if (builder.MoveToBookmark("InsertTableHere"))
        {
            // Build a 2x2 table at the bookmark location using the DocumentBuilder workflow.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 2,1");
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();
        }

        // Save the document to a file.
        const string outputPath = "TableAtBookmark.docx";
        doc.Save(outputPath);
    }
}
