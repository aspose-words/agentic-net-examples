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

        // Start the table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // ---------- Data rows ----------
        AddDataRow(builder, "Apples", "10", "1.20");
        AddDataRow(builder, "Bananas", "5", "0.80");
        AddDataRow(builder, "Carrots", "8", "0.50");

        // ---------- Footer row with totals ----------
        builder.InsertCell();
        builder.Write("Total");

        // Insert a field that sums the values above in the current column.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null);
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null);
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Update all fields so that the SUM fields are calculated.
        doc.UpdateFields();

        // Save the document to the local file system.
        doc.Save("TableWithFooter.docx");
    }

    // Helper method to add a data row to the table.
    private static void AddDataRow(DocumentBuilder builder, string item, string quantity, string price)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity);
        builder.InsertCell();
        builder.Write(price);
        builder.EndRow();
    }
}
