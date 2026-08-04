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

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndTable();

        // Give the table a fixed width so that text wrapping can be observed.
        table.PreferredWidth = PreferredWidth.FromPoints(200);

        // Set the table's text wrapping to "Around" (square style).
        table.TextWrapping = TextWrapping.Around;

        // Optional: define the distance between the table and surrounding text.
        table.AbsoluteHorizontalDistance = 20;
        table.AbsoluteVerticalDistance = 20;

        // Add a paragraph after the table to demonstrate the wrapping effect.
        builder.Writeln(
            "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
            "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Save the document to the local file system.
        doc.Save("TableWrapText.docx");
    }
}
