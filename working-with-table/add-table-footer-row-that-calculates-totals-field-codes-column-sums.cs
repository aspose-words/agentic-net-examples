using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableWithFooterExample
{
    static void Main()
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
        builder.EndRow();

        // ---------- Data rows ----------
        string[] items = { "Apples", "Bananas", "Carrots" };
        int[] quantities = { 20, 40, 50 };

        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);

            builder.InsertCell();
            builder.Write(quantities[i].ToString());

            builder.EndRow();
        }

        // ---------- Footer row (totals) ----------
        // First cell: label.
        builder.InsertCell();
        builder.Write("Total");

        // Second cell: formula field that sums all numbers above in this column.
        // InsertField expects a string for the field value (display text). Use null to let Word calculate the result.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null);

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Update all fields so the SUM field shows the calculated total.
        doc.UpdateFields();

        // Optional: Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("TableWithTotals.docx");
    }
}
