using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample XML data source.
        string xmlPath = Path.Combine(outputDir, "orders.xml");
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Orders>
  <Order>
    <OrderId>1001</OrderId>
    <CustomerName>John Doe</CustomerName>
    <LineItems>
      <LineItem>
        <ProductName>Widget A</ProductName>
        <Quantity>2</Quantity>
        <Price>9.99</Price>
      </LineItem>
      <LineItem>
        <ProductName>Widget B</ProductName>
        <Quantity>1</Quantity>
        <Price>19.99</Price>
      </LineItem>
    </LineItems>
  </Order>
  <Order>
    <OrderId>1002</OrderId>
    <CustomerName>Jane Smith</CustomerName>
    <LineItems>
      <LineItem>
        <ProductName>Gadget X</ProductName>
        <Quantity>5</Quantity>
        <Price>4.99</Price>
      </LineItem>
    </LineItems>
  </Order>
</Orders>");

        // Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("------------------------------");
        builder.Writeln("<<foreach [order in Orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Line Items:");
        builder.Writeln("<<foreach [item in order.LineItems.LineItem]>>");
        builder.Writeln("  - Product: <<[item.ProductName]>>");
        builder.Writeln("    Quantity: <<[item.Quantity]>>");
        builder.Writeln("    Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("------------------------------");

        templateDoc.Save(templatePath);

        // Load the template.
        Document reportDoc = new(templatePath);

        // Load XML data source.
        XmlDataSource xmlData = new(xmlPath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, xmlData, "Orders");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {reportPath}");
    }
}
