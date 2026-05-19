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

        // Start a table and add a few rows/cells.
        Table table = builder.StartTable();

        // First row (header)
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Third row
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the whole table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;
        // Apply the style to all parts of the table (header, rows, columns, banding, etc.).
        table.StyleOptions = TableStyleOptions.Default;

        // Optionally auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the local file system.
        doc.Save("GridTableStyle.docx");
    }
}
