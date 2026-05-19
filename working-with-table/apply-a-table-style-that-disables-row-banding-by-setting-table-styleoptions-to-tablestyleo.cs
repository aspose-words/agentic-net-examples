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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style to the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Disable row banding by clearing all style options.
        // TableStyleOptions.RowBands is the flag that enables banding; removing it (or using None) disables it.
        table.StyleOptions = TableStyleOptions.None;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableNoRowBanding.docx");
        doc.Save(outputPath);
    }
}
