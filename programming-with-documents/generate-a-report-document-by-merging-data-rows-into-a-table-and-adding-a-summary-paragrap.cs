using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    // Simple data model for report rows
    public class ReportItem
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }

        public decimal Total => Quantity * UnitPrice;
    }

    public static void Main()
    {
        // Sample data to merge into the report
        List<ReportItem> items = new List<ReportItem>
        {
            new ReportItem { Product = "Apple",  Quantity = 10, UnitPrice = 0.5m },
            new ReportItem { Product = "Banana", Quantity = 5,  UnitPrice = 0.3m },
            new ReportItem { Product = "Orange", Quantity = 8,  UnitPrice = 0.4m }
        };

        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a title
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.Font.Size = 16;
        builder.Font.Bold = true;
        builder.Writeln("Sales Report");
        builder.Font.ClearFormatting();
        builder.Writeln(); // empty line

        // Insert a table with 4 columns: Product, Quantity, Unit Price, Total
        Table table = builder.StartTable();

        // Define column widths (optional)
        builder.InsertCell();
        builder.CellFormat.Width = 150;
        builder.Font.Bold = true;
        builder.Writeln("Product");

        builder.InsertCell();
        builder.CellFormat.Width = 80;
        builder.Writeln("Quantity");

        builder.InsertCell();
        builder.CellFormat.Width = 80;
        builder.Writeln("Unit Price");

        builder.InsertCell();
        builder.CellFormat.Width = 80;
        builder.Writeln("Total");

        // End header row
        builder.EndRow();

        // Populate table rows with data
        builder.Font.Bold = false;
        foreach (var item in items)
        {
            builder.InsertCell();
            builder.Writeln(item.Product);

            builder.InsertCell();
            builder.Writeln(item.Quantity.ToString());

            builder.InsertCell();
            builder.Writeln(item.UnitPrice.ToString("C"));

            builder.InsertCell();
            builder.Writeln(item.Total.ToString("C"));

            builder.EndRow();
        }

        // End the table
        builder.EndTable();

        // Add a summary paragraph below the table
        builder.Writeln(); // empty line
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
        builder.Font.Size = 12;
        builder.Font.Bold = true;

        int totalQuantity = 0;
        decimal grandTotal = 0m;
        foreach (var item in items)
        {
            totalQuantity += item.Quantity;
            grandTotal += item.Total;
        }

        builder.Writeln($"Total Items: {totalQuantity}");
        builder.Writeln($"Grand Total: {grandTotal:C}");

        // Save the document
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Verify that the file was created and can be reopened
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // Optionally, you could perform further checks here.
        }
    }
}
