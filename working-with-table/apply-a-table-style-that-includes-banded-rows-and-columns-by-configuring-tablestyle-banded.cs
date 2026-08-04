using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. The first call to InsertCell creates the first row.
        Table table = builder.StartTable();
        builder.InsertCell(); // first cell of the first row

        // Apply a built‑in table style that supports banding.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable both row banding and column banding.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.ColumnBands;

        // Let the table size itself to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Fill the first row (header).
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.InsertCell();
        builder.Writeln("Header 3");
        builder.EndRow();

        // Add two more rows with sample data.
        for (int row = 2; row <= 3; row++)
        {
            builder.InsertCell();
            builder.Writeln($"R{row}C1");
            builder.InsertCell();
            builder.Writeln($"R{row}C2");
            builder.InsertCell();
            builder.Writeln($"R{row}C3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableWithBandedRowsAndColumns.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
