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

        // Build a simple table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header");
        builder.EndRow();

        // Second row (data).
        builder.InsertCell();
        builder.Write("Data");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific style options using the TableStyleOptions flags.
        // Here we enable formatting for the first row, first column, and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the document to a DOCX file.
        doc.Save("TableWithStyleOptions.docx");
    }
}
