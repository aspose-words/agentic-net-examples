using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code pages provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "customers.json";
        var sampleData = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer
                {
                    Id = 1,
                    Name = "Alice Johnson",
                    Email = "alice@example.com",
                    Orders = new List<Order>
                    {
                        new Order { OrderId = 101, Product = "Laptop", Amount = 1299.99m },
                        new Order { OrderId = 102, Product = "Mouse", Amount = 25.50m }
                    }
                },
                new Customer
                {
                    Id = 2,
                    Name = "Bob Smith",
                    Email = "bob@example.com",
                    Orders = new List<Order>
                    {
                        new Order { OrderId = 201, Product = "Smartphone", Amount = 799.00m }
                    }
                }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Customer Report");
        builder.Writeln("");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name: <<[c.Name]>>");
        builder.Writeln("Email: <<[c.Email]>>");
        builder.Writeln("Orders:");
        builder.Writeln("<<foreach [o in c.Orders]>>");
        builder.Writeln("- Product: <<[o.Product]>>, Amount: $<<[o.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load JSON data.
        string json = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<ReportModel>(json)!;

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "CustomerReport.docx";
        reportDoc.Save(outputPath);
    }
}

public class ReportModel
{
    public List<Customer> Customers { get; set; } = new();
}

public class Customer
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Email { get; set; } = "";
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int OrderId { get; set; }
    public string Product { get; set; } = "";
    public decimal Amount { get; set; }
}
