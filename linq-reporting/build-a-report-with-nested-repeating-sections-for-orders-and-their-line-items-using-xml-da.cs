using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFile = "Orders.xml";
        File.WriteAllText(xmlFile, GetSampleXml());

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach – iterate over orders.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order Date: <<[order.OrderDate]>>");
        builder.Writeln("Items:");
        // Inner foreach – iterate over line items of the current order.
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        // Correct arithmetic expression syntax: the whole expression must be inside a single <<[ ... ]>> tag.
        builder.Writeln("- <<[item.ProductName]>>: <<[item.Quantity]>> x <<[item.Price]>> = <<[item.Quantity * item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the XML data source.
        ReportingEngine engine = new ReportingEngine();
        XmlDataSource dataSource = new XmlDataSource(xmlFile);
        engine.BuildReport(template, dataSource, "orders");

        // Save the generated report.
        template.Save("OrdersReport.docx");
    }

    // Returns a simple XML string containing two orders with line items.
    private static string GetSampleXml()
    {
        return @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
  <Order>
    <CustomerName>John Doe</CustomerName>
    <OrderDate>2023-08-01</OrderDate>
    <Items>
      <Item>
        <ProductName>Widget A</ProductName>
        <Quantity>2</Quantity>
        <Price>9.99</Price>
      </Item>
      <Item>
        <ProductName>Gadget B</ProductName>
        <Quantity>1</Quantity>
        <Price>19.95</Price>
      </Item>
    </Items>
  </Order>
  <Order>
    <CustomerName>Jane Smith</CustomerName>
    <OrderDate>2023-08-03</OrderDate>
    <Items>
      <Item>
        <ProductName>Thingamajig</ProductName>
        <Quantity>5</Quantity>
        <Price>3.50</Price>
      </Item>
    </Items>
  </Order>
</Orders>";
    }
}
