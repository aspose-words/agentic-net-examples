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

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set a preferred width so the table is visible.
        table.PreferredWidth = PreferredWidth.FromPoints(300);

        // Configure the table to wrap text around it.
        table.TextWrapping = TextWrapping.Around;
        // Optional: set distances from surrounding text.
        table.AbsoluteHorizontalDistance = 20;
        table.AbsoluteVerticalDistance = 10;

        // Add some surrounding text after the table.
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document to the local file system.
        doc.Save("TableWrapAround.docx");
    }
}
