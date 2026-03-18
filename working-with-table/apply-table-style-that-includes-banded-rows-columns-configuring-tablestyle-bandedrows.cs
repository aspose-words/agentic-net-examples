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

        // Insert a header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Insert a few data rows.
        for (int i = 0; i < 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} Col 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable both row banding and column banding.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.ColumnBands;

        // Optionally also apply first row/column formatting.
        table.StyleOptions |= TableStyleOptions.FirstRow | TableStyleOptions.FirstColumn;

        // Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("TableWithBandedRowsAndColumns.docx");
    }
}
