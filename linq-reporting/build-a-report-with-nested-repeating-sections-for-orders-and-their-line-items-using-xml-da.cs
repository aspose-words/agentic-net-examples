using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            string xmlPath = Path.Combine(outputDir, "orders.xml");
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");

            // -----------------------------------------------------------------
            // 1. Create a sample XML data source with orders and line items.
            // -----------------------------------------------------------------
            File.WriteAllText(xmlPath,
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderId>1001</OrderId>
        <Items>
            <Item>
                <ProductName>Apple</ProductName>
                <Quantity>3</Quantity>
                <Price>0.5</Price>
            </Item>
            <Item>
                <ProductName>Banana</ProductName>
                <Quantity>5</Quantity>
                <Price>0.3</Price>
            </Item>
        </Items>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderId>1002</OrderId>
        <Items>
            <Item>
                <ProductName>Orange</ProductName>
                <Quantity>2</Quantity>
                <Price>0.7</Price>
            </Item>
            <Item>
                <ProductName>Grapes</ProductName>
                <Quantity>1</Quantity>
                <Price>2.0</Price>
            </Item>
        </Items>
    </Order>
</Orders>");

            // -----------------------------------------------------------------
            // 2. Build a template document programmatically with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Outer foreach over orders.
            builder.Writeln("<<foreach [order in orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln("Order ID: <<[order.OrderId]>>");
            builder.Writeln();

            // Header for line items.
            builder.Writeln("Items:");
            // Inner foreach over items.
            builder.Writeln("<<foreach [item in order.Items.Item]>>");
            builder.Writeln("- Product: <<[item.ProductName]>>");
            builder.Writeln("  Quantity: <<[item.Quantity]>>");
            builder.Writeln("  Price: $<<[item.Price]>>");
            builder.Writeln("<</foreach>>"); // End inner foreach

            builder.Writeln("<</foreach>>"); // End outer foreach

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and generate the report using the XML data source.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options required.

            // The root name used in the template tags is "orders".
            engine.BuildReport(reportDoc, xmlDataSource, "orders");

            // Save the generated report.
            reportDoc.Save(reportPath);

            Console.WriteLine($"Report generated at: {reportPath}");
        }
    }
}
