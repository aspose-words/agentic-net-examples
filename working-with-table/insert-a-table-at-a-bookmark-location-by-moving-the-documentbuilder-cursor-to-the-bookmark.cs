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

        // Add a bookmark that will mark the insertion point for the table.
        builder.StartBookmark("TableBookmark");
        builder.Writeln("Text before the table.");
        builder.EndBookmark("TableBookmark");

        // Move the builder cursor to the start of the bookmark (after the bookmark start node).
        // This positions the insertion point exactly where the table should be placed.
        builder.MoveToBookmark("TableBookmark", isStart: true, isAfter: true);

        // Build the table at the bookmark location using the DocumentBuilder workflow.
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

        // Save the resulting document.
        doc.Save("TableAtBookmark.docx");
    }
}
