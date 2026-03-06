using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to construct its contents.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. At least one cell must be inserted before any formatting.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style using its identifier.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Resize the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("TableWithStyle.docx");
    }
}
