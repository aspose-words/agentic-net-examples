using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data
        string xmlPath = "orders.xml";
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <Id>1</Id>
        <Customer>John Doe</Customer>
        <Amount>150.00</Amount>
    </Order>
    <Order>
        <Id>2</Id>
        <Customer>Jane Smith</Customer>
        <Amount>230.50</Amount>
    </Order>
    <Order>
        <Id>3</Id>
        <Customer>Bob Johnson</Customer>
        <Amount>99.99</Amount>
    </Order>
</Orders>");

        // Load XML and map to model
        XDocument doc = XDocument.Load(xmlPath);
        ReportModel model = new()
        {
            Orders = doc.Root?
                .Elements("Order")
                .Select(o => new Order
                {
                    Id = (string?)o.Element("Id") ?? string.Empty,
                    Customer = (string?)o.Element("Customer") ?? string.Empty,
                    Amount = (string?)o.Element("Amount") ?? string.Empty
                })
                .ToList() ?? new()
        };

        // Create template document programmatically
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Begin foreach over Orders
        builder.Writeln("<<foreach [order in Orders]>>");

        // Start table
        Table table = builder.StartTable();

        // Header row (element names)
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Customer");
        builder.InsertCell();
        builder.Writeln("Amount");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[order.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Customer]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Amount]>>");
        builder.EndRow();

        // End table
        builder.EndTable();

        // End foreach
        builder.Writeln("<</foreach>>");

        // Save template
        templateDoc.Save(templatePath);

        // Load template for reporting
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save final report
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Wrapper model for the report
public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

// Simple order data class
public class Order
{
    public string Id { get; set; } = string.Empty;
    public string Customer { get; set; } = string.Empty;
    public string Amount { get; set; } = string.Empty;
}
