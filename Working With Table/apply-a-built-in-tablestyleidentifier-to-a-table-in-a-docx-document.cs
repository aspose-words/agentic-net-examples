using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert at least one cell (required before setting table formatting).
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // Finish the first row and the table.
        builder.EndRow();
        builder.EndTable();

        // Apply a built‑in table style by its identifier.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document to a DOCX file.
        doc.Save("TableWithStyle.docx");
    }
}
