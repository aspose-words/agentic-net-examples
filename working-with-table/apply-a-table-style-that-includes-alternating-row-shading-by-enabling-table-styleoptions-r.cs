using System;
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
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Sample data rows.
        string[,] data = {
            { "Apples",  "10" },
            { "Bananas", "20" },
            { "Carrots", "30" },
            { "Dates",   "40" }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            builder.InsertCell();
            builder.Write(data[i, 0]);
            builder.InsertCell();
            builder.Write(data[i, 1]);
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in style that supports shading.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable row banding (alternating row shading) and keep the first row styled as a header.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.FirstRow;

        // Save the document to the local file system.
        doc.Save("TableWithRowBanding.docx");
    }
}
