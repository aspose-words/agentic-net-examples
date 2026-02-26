using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple 2x2 table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Optionally assign a built‑in style to the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific TableStyleOptions flags (e.g., first row and row banding).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document as an RTF file.
        doc.Save("TableWithStyleOptions.rtf");
    }
}
