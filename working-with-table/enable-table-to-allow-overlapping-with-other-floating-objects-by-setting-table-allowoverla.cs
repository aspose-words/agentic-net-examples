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

        // Start building a table.
        Table table = builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Write("Floating table cell.");

        // Finish the table.
        builder.EndTable();

        // Make the table a floating object by enabling text wrapping.
        table.TextWrapping = TextWrapping.Around;

        // Position the floating table on the page.
        table.AbsoluteHorizontalDistance = 50; // points from the anchor.
        table.AbsoluteVerticalDistance = 50;   // points from the anchor.

        // No explicit validation needed – AllowOverlap is true by default for floating tables.

        // Save the document to a file.
        doc.Save("FloatingTableAllowOverlap.docx");
    }
}
