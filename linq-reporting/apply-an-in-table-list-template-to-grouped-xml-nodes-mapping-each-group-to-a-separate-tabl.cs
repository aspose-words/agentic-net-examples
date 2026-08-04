using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        Directory.CreateDirectory("output");

        // 1. Create sample XML data
        const string xmlPath = "data.xml";
        File.WriteAllText(xmlPath,
@"<Orders>
    <Order>
        <Id>1</Id>
        <CustomerName>Acme Corp</CustomerName>
        <ProductName>Widget A</ProductName>
        <Price>19.99</Price>
    </Order>
    <Order>
        <Id>2</Id>
        <CustomerName>Acme Corp</CustomerName>
        <ProductName>Widget B</ProductName>
        <Price>29.99</Price>
    </Order>
    <Order>
        <Id>3</Id>
        <CustomerName>Beta Ltd</CustomerName>
        <ProductName>Gadget X</ProductName>
        <Price>49.50</Price>
    </Order>
    <Order>
        <Id>4</Id>
        <CustomerName>Beta Ltd</CustomerName>
        <ProductName>Gadget Y</ProductName>
        <Price>59.75</Price>
    </Order>
    <Order>
        <Id>5</Id>
        <CustomerName>Gamma Inc</CustomerName>
        <ProductName>Thingamajig</ProductName>
        <Price>99.00</Price>
    </Order>
</Orders>");

        // 2. Load XML and transform into grouped model
        XDocument xDoc = XDocument.Load(xmlPath);
        var flatOrders = xDoc.Root!
            .Elements("Order")
            .Select(o => new Order
            {
                Id = (int)o.Element("Id")!,
                CustomerName = (string)o.Element("CustomerName")!,
                ProductName = (string)o.Element("ProductName")!,
                Price = (decimal)o.Element("Price")!
            })
            .ToList();

        var groups = flatOrders
            .GroupBy(o => o.CustomerName)
            .Select(g => new CustomerGroup
            {
                CustomerName = g.Key,
                Orders = g.Select(o => new Order
                {
                    Id = o.Id,
                    ProductName = o.ProductName,
                    Price = o.Price
                }).ToList()
            })
            .ToList();

        var model = new ReportModel { Groups = groups };

        // 3. Build the LINQ Reporting template programmatically
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach over groups
        builder.Writeln("<<foreach [group in Model.Groups]>>");
        builder.Writeln("Customer: <<[group.CustomerName]>>");
        builder.Writeln();

        // Header table for each group
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Order Id");
        builder.InsertCell();
        builder.Writeln("Product");
        builder.InsertCell();
        builder.Writeln("Price");
        builder.EndRow();
        builder.EndTable();

        // Inner foreach over orders – each order gets its own row table
        builder.Writeln("<<foreach [order in group.Orders]>>");
        Table orderTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[order.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.ProductName]>>");
        builder.InsertCell();
        builder.Writeln("<<[order.Price]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Blank line between groups
        builder.Writeln();

        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // 4. Load the template and generate the report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "Model");

        // 5. Save the final report
        const string outputPath = "output/OrdersReport.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Data model classes
public class ReportModel
{
    public List<CustomerGroup> Groups { get; set; } = new();
}

public class CustomerGroup
{
    public string CustomerName { get; set; } = string.Empty;
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public int Id { get; set; }
    public string ProductName { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public string CustomerName { get; set; } = string.Empty;
}
