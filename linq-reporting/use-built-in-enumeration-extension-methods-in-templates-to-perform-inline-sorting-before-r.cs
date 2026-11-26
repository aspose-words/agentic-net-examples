using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data with unsorted products.
        var order = new Order
        {
            CustomerName = "John Doe",
            Products = new List<Product>
            {
                new Product { Name = "Keyboard", Price = 49.99m },
                new Product { Name = "Monitor", Price = 199.99m },
                new Product { Name = "Mouse", Price = 19.99m },
                new Product { Name = "Laptop", Price = 999.99m }
            }
        };

        // Create the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Products sorted by price (ascending):");
        builder.Writeln("<<foreach [p in order.Products.OrderBy(p => p.Price)]>>");
        builder.Writeln("- <<[p.Name]>>: $<<[p.Price]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, order, "order");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine(success ? $"Report generated: {reportPath}" : "Report generation failed.");
    }
}

// Public data model classes.
public class Order
{
    public string CustomerName { get; set; } = "";
    public List<Product> Products { get; set; } = new();
}

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
}
