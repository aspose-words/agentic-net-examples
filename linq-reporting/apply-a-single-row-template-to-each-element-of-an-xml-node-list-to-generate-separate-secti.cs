using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Order
{
    public string Id { get; set; } = "";
    public string Customer { get; set; } = "";
}

public class OrdersRoot
{
    public List<Order> Order { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlPath = "orders.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <Id>1</Id>
        <Customer>John Doe</Customer>
    </Order>
    <Order>
        <Id>2</Id>
        <Customer>Jane Smith</Customer>
    </Order>
    <Order>
        <Id>3</Id>
        <Customer>Bob Johnson</Customer>
    </Order>
</Orders>";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML into the strongly‑typed model.
        OrdersRoot data = new()
        {
            Order = new List<Order>()
        };
        var xmlDoc = new System.Xml.XmlDocument();
        xmlDoc.Load(xmlPath);
        var orderNodes = xmlDoc.SelectNodes("/Orders/Order");
        if (orderNodes != null)
        {
            foreach (System.Xml.XmlNode node in orderNodes)
            {
                data.Order.Add(new Order
                {
                    Id = node.SelectSingleNode("Id")?.InnerText ?? "",
                    Customer = node.SelectSingleNode("Customer")?.InnerText ?? ""
                });
            }
        }

        // Create the LINQ Reporting template.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("Orders Report");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Foreach block over the Order collection.
        builder.Writeln("<<foreach [order in Order]>>");
        builder.Writeln("Order ID: <<[order.Id]>>");
        builder.Writeln("Customer: <<[order.Customer]>>");
        // Separate each order with a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the strongly‑typed data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, data, "orders");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
