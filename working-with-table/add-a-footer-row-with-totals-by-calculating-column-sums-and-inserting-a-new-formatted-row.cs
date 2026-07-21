using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Path for the output document.
        string outputPath = "TableWithFooter.docx";

        // Create a new blank document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // ---------- Data rows ----------
        AddDataRow(builder, "Apples", 20);
        AddDataRow(builder, "Bananas", 40);
        AddDataRow(builder, "Carrots", 50);

        // ---------- Calculate column sums ----------
        double quantitySum = 0;
        // Skip the header row (index 0).
        for (int i = 1; i < table.Rows.Count; i++)
        {
            Row row = table.Rows[i];
            // The quantity is in the second cell (index 1).
            string cellText = row.Cells[1].ToString(SaveFormat.Text).Trim();
            if (double.TryParse(cellText, out double value))
                quantitySum += value;
        }

        // ---------- Footer row with totals ----------
        // Apply formatting only to this row.
        builder.Font.Bold = true;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        builder.Write(quantitySum.ToString());
        builder.EndRow();

        // Reset formatting so it does not affect later content.
        builder.Font.Bold = false;
        builder.CellFormat.Shading.ClearFormatting();

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");
    }

    // Helper method to add a data row with an item name and a numeric quantity.
    private static void AddDataRow(DocumentBuilder builder, string item, int quantity)
    {
        builder.InsertCell();
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity.ToString());
        builder.EndRow();
    }
}
