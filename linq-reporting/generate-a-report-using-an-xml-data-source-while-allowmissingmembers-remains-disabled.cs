using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define file paths
        string currentDir = Directory.GetCurrentDirectory();
        string xmlPath = Path.Combine(currentDir, "sample.xml");
        string templatePath = Path.Combine(currentDir, "Template.docx");
        string reportPath = Path.Combine(currentDir, "Report.docx");

        // Create sample XML data
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <Total>123.45</Total>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <Total>678.90</Total>
    </Order>
</Orders>";
        File.WriteAllText(xmlPath, xmlContent);

        // Deserialize XML into a strongly‑typed model
        OrdersRoot dataModel;
        var serializer = new XmlSerializer(typeof(OrdersRoot));
        using (var fs = new FileStream(xmlPath, FileMode.Open, FileAccess.Read))
        {
            dataModel = (OrdersRoot)serializer.Deserialize(fs)!;
        }

        // Build the template document with LINQ Reporting tags
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Order Report");
        builder.Writeln("<<foreach [order in data.Orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>   Total: <<[order.Total]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Build the report (AllowMissingMembers remains disabled)
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, dataModel, "data");

        // Save the final report
        reportDoc.Save(reportPath);
    }
}

// Strongly‑typed model that matches the XML structure
[XmlRoot("Orders")]
public class OrdersRoot
{
    [XmlElement("Order")]
    public List<Order> Orders { get; set; } = new();
}

public class Order
{
    public string CustomerName { get; set; } = "";
    public decimal Total { get; set; }
}
