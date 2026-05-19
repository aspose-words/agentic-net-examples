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

        // First row – header.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("40");
        builder.EndRow();

        // Fourth row.
        builder.InsertCell();
        builder.Writeln("Carrots");
        builder.InsertCell();
        builder.Writeln("50");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style that supports shading.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable row banding (alternating row shading).
        table.StyleOptions = TableStyleOptions.RowBands;

        // Adjust the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the local file system.
        doc.Save("TableWithRowBanding.docx");
    }
}
