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

        // Start a table. The first InsertCell call is required before any table formatting.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Apply a built‑in style and enable row banding (alternating row shading).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
        table.StyleOptions = TableStyleOptions.RowBands;

        // Optional: let the table auto‑fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Populate the table with sample data (two columns, three rows).
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Table.RowBanding.docx");
        doc.Save(outputPath);
    }
}
