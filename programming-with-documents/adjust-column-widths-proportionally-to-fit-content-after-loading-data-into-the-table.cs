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

        // Start a table and keep a reference to it.
        Table table = builder.StartTable();

        // Populate the table with sample data.
        // Header row.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Description");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apple");
        builder.InsertCell();
        builder.Write("Fresh red apples from the orchard");
        builder.InsertCell();
        builder.Write("$1.20");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Banana");
        builder.InsertCell();
        builder.Write("Ripe bananas, sweet and soft");
        builder.InsertCell();
        builder.Write("$0.80");
        builder.EndRow();

        // Third data row.
        builder.InsertCell();
        builder.Write("Cherry");
        builder.InsertCell();
        builder.Write("Organic cherries, packed in a box");
        builder.InsertCell();
        builder.Write("$3.50");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Adjust column widths proportionally to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the local file system.
        doc.Save("AdjustedTable.docx");
    }
}
