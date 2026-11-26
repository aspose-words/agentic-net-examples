using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string CustomerName { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}

public class ReportModel
{
    public List<Order> Orders { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Create sample XML data file (data.xml)
        // -----------------------------------------------------------------
        string xmlPath = "data.xml";
        var xmlDoc = new XDocument(
            new XElement("Orders",
                new XElement("Order",
                    new XElement("CustomerName", "John Doe"),
                    new XElement("Amount", "123.45")),
                new XElement("Order",
                    new XElement("CustomerName", "Jane Smith"),
                    new XElement("Amount", "678.90"))
            )
        );
        xmlDoc.Save(xmlPath);

        // -----------------------------------------------------------------
        // Load XML data into strongly‑typed model
        // -----------------------------------------------------------------
        var model = new ReportModel();
        foreach (var orderElem in XDocument.Load(xmlPath).Root?.Elements("Order") ?? [])
        {
            var order = new Order
            {
                CustomerName = (string?)orderElem.Element("CustomerName") ?? string.Empty,
                Amount = decimal.TryParse((string?)orderElem.Element("Amount"), out var amt) ? amt : 0m
            };
            model.Orders.Add(order);
        }

        // -----------------------------------------------------------------
        // Create a template document with LINQ Reporting tags (template.docx)
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Amount: <<[order.Amount]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template document
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // Build the report using the strongly‑typed model
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
