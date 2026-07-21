using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell (required before setting table formatting).
        builder.InsertCell();

        // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;
        // Apply the style to all parts of the table (rows, columns, bands, etc.).
        table.StyleOptions = TableStyleOptions.Default;

        // Populate the table with some sample data.
        // Header row.
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("35");
        builder.EndRow();

        // Third data row.
        builder.InsertCell();
        builder.Writeln("Carrots");
        builder.InsertCell();
        builder.Writeln("50");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "TableWithGridTable5Dark.docx");
        doc.Save(outputPath);
    }
}
