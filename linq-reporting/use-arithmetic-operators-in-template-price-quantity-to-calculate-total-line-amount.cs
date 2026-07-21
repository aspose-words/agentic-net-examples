using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class OrderItem
{
    public string Description { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}

public class Order
{
    public List<OrderItem> Items { get; set; } = new();
}

public partial class Program
{
    public static void Main()
    {
        // 1. Create the LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Invoice");
        builder.Writeln(); // Empty line.

        // Begin a foreach loop over the collection Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Output each field and calculate the line total using arithmetic operators.
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("Price: $<<[item.Price]>>");
        builder.Writeln("Quantity: <<[item.Quantity]>>");
        builder.Writeln("Line Total: $<<[item.Price] * [item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // 2. Prepare sample data.
        Order order = new Order();
        order.Items.Add(new OrderItem
        {
            Description = "Widget",
            Price = 9.99m,
            Quantity = 3
        });
        order.Items.Add(new OrderItem
        {
            Description = "Gadget",
            Price = 14.50m,
            Quantity = 2
        });

        // 3. Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // The root object name used in the template is "order".
        bool success = engine.BuildReport(reportDoc, order, "order");

        // 4. Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
