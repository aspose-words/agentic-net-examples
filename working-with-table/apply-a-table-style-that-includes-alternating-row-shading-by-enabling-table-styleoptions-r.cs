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

        // Header row – required before applying any style settings.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Apply a built‑in style that supports row banding.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable only row banding (alternating row shading).
        table.StyleOptions = TableStyleOptions.RowBands;

        // Add a few data rows.
        for (int i = 1; i <= 4; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln($"{i * 10}");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        doc.Save("TableStyleRowBanding.docx");
    }
}
