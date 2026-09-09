using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data source file.
        const string xmlFileName = "Orders.xml";
        File.WriteAllText(xmlFileName,
@"<Orders>
    <Order>
        <OrderId>1001</OrderId>
        <CustomerName>John Doe</CustomerName>
        <Items>
            <Item>
                <ProductName>Widget A</ProductName>
                <Quantity>3</Quantity>
            </Item>
            <Item>
                <ProductName>Gadget B</ProductName>
                <Quantity>1</Quantity>
            </Item>
        </Items>
    </Order>
    <Order>
        <OrderId>1002</OrderId>
        <CustomerName>Jane Smith</CustomerName>
        <Items>
            <Item>
                <ProductName>Widget C</ProductName>
                <Quantity>2</Quantity>
            </Item>
        </Items>
    </Order>
</Orders>");

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach – iterate over orders.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        // Inner foreach – iterate over line items of the current order.
        // For XML data source the collection of items is accessed via order.Items.Item
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.ProductName]>> x <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlFileName);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "orders");

        // Save the generated report.
        const string outputFileName = "OrdersReport.docx";
        template.Save(outputFileName);
    }
}
