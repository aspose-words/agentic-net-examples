using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data.
        string xmlPath = "orders.xml";
        File.WriteAllText(xmlPath,
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderId>1001</OrderId>
        <Items>
            <Item>
                <Product>Apple</Product>
                <Quantity>2</Quantity>
            </Item>
            <Item>
                <Product>Banana</Product>
                <Quantity>5</Quantity>
            </Item>
        </Items>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderId>1002</OrderId>
        <Items>
            <Item>
                <Product>Orange</Product>
                <Quantity>3</Quantity>
            </Item>
        </Items>
    </Order>
</Orders>");

        // Build a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln();

        // The XML data source is named "orders".
        // Since the root element contains a collection of <Order> elements,
        // we iterate directly over the root collection.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Items:");
        // Inner foreach – iterate over Item elements within the current order.
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.Product]>> x <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Load XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report.
        var engine = new ReportingEngine { Options = ReportBuildOptions.None };
        // Use the name "orders" to reference the data source in the template.
        engine.BuildReport(reportDoc, xmlDataSource, "orders");

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
