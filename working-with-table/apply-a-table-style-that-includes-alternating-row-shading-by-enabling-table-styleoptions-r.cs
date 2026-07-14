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

        // Start building a table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity (kg)");
        builder.EndRow();

        // Data rows.
        AddRow(builder, "Apples", "20");
        AddRow(builder, "Bananas", "40");
        AddRow(builder, "Carrots", "50");

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style that supports row banding.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable alternating row shading (row banding).
        table.StyleOptions = TableStyleOptions.RowBands;

        // Adjust the table size to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TableWithRowBanding.docx");
        doc.Save(outputPath);
    }

    private static void AddRow(DocumentBuilder builder, string item, string quantity)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity);
        builder.EndRow();
    }
}
