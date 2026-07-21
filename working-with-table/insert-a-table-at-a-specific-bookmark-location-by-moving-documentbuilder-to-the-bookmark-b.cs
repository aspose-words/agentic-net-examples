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

        // Add some initial content.
        builder.Writeln("Document start.");

        // Create a bookmark that will mark the insertion point for the table.
        builder.StartBookmark("MyTableBookmark");
        builder.Writeln("Placeholder for table.");
        builder.EndBookmark("MyTableBookmark");

        // Move the builder to the bookmark. This positions the cursor just after the bookmark start,
        // which is a valid location for inserting a block-level node such as a table.
        builder.MoveToBookmark("MyTableBookmark");

        // Build the table at the bookmark location.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the resulting document to the current directory.
        doc.Save("TableAtBookmark.docx");
    }
}
