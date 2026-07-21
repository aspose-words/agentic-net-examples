using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Step 1: Create the DOCX template programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Step 2: Load the template.
        var doc = new Document(templatePath);

        // Step 3: Prepare sample data.
        var order = new Order
        {
            CustomerName = "John Doe",
            Items = new List<OrderItem>
            {
                new OrderItem { Index = 1, Name = "Laptop", Price = 1200.00 },
                new OrderItem { Index = 2, Name = "Mouse", Price = 25.50 },
                new OrderItem { Index = 3, Name = "Keyboard", Price = 45.00 }
            }
        };

        // Step 4: Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, order, "order");

        // Step 5: Save the populated document as PDF.
        doc.Save("Report.pdf", SaveFormat.Pdf);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Order Report");
        builder.Writeln();

        // Customer name tag
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln();

        // Begin foreach loop over Items
        builder.Writeln("<<foreach [item in Items]>>");

        // Table with header row
        var table = builder.StartTable();

        // Header cells
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();

        // Data row (repeated for each item)
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();

        // Price with conditional background color for expensive items (> 100)
        builder.Writeln(
            "<<if [item.Price > 100]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Price]>> <</backColor>><</if>>" +
            "<<if [item.Price <= 100]>> <<[item.Price]>> <</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(filePath);
    }
}

// Data model classes
public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public List<OrderItem> Items { get; set; } = new();
}

public class OrderItem
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
}
