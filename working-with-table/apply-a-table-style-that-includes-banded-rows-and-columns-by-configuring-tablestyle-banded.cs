using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a simple 2‑column table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add a few data rows.
        for (int i = 1; i <= 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style that supports banding.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;

        // Enable both row banding and column banding.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.ColumnBands;

        // Resize the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Ensure the output directory exists and save the document.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWithBandedRowsColumns.docx");
        doc.Save(outputPath);
    }
}
