#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using System.Xml.Linq;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible XML encoding issues.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample XML data.
        string xmlContent = @"
<Orders>
    <Order>
        <Id>1</Id>
        <Customer>John Doe</Customer>
        <Amount>100.00</Amount>
    </Order>
    <Order>
        <Id>2</Id>
        <Customer>Jane Smith</Customer>
        <Amount>150.50</Amount>
    </Order>
    <Order>
        <Id>3</Id>
        <Customer>Bob Johnson</Customer>
        <Amount>200.75</Amount>
    </Order>
</Orders>";

        // Parse XML into model objects.
        XDocument xDoc = XDocument.Parse(xmlContent);
        List<Order> orders = xDoc.Root?.Elements("Order")
            .Select(o => new Order
            {
                Id = (string?)o.Element("Id") ?? string.Empty,
                Customer = (string?)o.Element("Customer") ?? string.Empty,
                Amount = (string?)o.Element("Amount") ?? string.Empty
            })
            .ToList() ?? new List<Order>();

        var model = new OrderData { Orders = orders };

        // Create template document.
        const string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin foreach block.
        builder.Writeln("<<foreach [order in Orders]>>");

        // Start table.
        Table table = builder.StartTable();

        // Header row – use property names of Order class.
        PropertyInfo[] props = typeof(Order).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        foreach (PropertyInfo prop in props)
        {
            builder.InsertCell();
            builder.Writeln(prop.Name);
        }
        builder.EndRow();

        // Data row – placeholders for each property.
        foreach (PropertyInfo prop in props)
        {
            builder.InsertCell();
            builder.Writeln($"<<[order.{prop.Name}]>>");
        }
        builder.EndRow();

        // End table.
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load template for reporting.
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save output.
        string outputPath = Path.Combine("output", "Report.docx");
        Directory.CreateDirectory("output");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}

// Wrapper model.
public class OrderData
{
    public List<Order> Orders { get; set; } = new();
}

// Individual order.
public class Order
{
    public string Id { get; set; } = string.Empty;
    public string Customer { get; set; } = string.Empty;
    public string Amount { get; set; } = string.Empty;
}
