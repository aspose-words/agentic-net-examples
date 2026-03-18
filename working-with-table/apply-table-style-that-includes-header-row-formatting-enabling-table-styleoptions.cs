using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleWithHeader
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first cell to satisfy the requirement of having at least one row before formatting.
        builder.InsertCell();

        // Choose a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable the first‑row conditional formatting (header row) for the chosen style.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Build the header row.
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Add a few data rows.
        builder.InsertCell();
        builder.Writeln("Apple");
        builder.InsertCell();
        builder.Writeln("$1.00");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Banana");
        builder.InsertCell();
        builder.Writeln("$0.50");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        doc.Save("TableWithHeaderStyle.docx");
    }
}
