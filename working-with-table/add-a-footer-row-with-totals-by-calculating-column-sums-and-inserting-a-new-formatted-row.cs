using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // ---------- Header Row ----------
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Sample data rows.
        double totalQuantity = 0;
        string[] items = { "Apples", "Bananas", "Carrots", "Dates" };
        double[] quantities = { 20, 35, 50, 15 };

        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);

            builder.InsertCell();
            builder.Write(quantities[i].ToString());

            builder.EndRow();

            totalQuantity += quantities[i];
        }

        // ---------- Footer Row (Totals) ----------
        // Apply formatting for the footer row.
        builder.RowFormat.Height = 20;
        builder.RowFormat.HeightRule = HeightRule.AtLeast;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Font.Bold = true;

        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        builder.Write(totalQuantity.ToString());

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFooter.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program finishes here without waiting for user input.
    }
}
