using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template, XML data source and the final report.
        string templatePath = "Template.docx";
        string xmlDataPath = "Orders.xml";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer foreach – iterate over each Order element.
        // The XML data source is named "orders", which represents a collection of Order rows.
        builder.Writeln("<<foreach [order in orders]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Order ID: <<[order.OrderId]>>");
        builder.Writeln("Items:");
        // Inner foreach – iterate over each Item within the current Order.
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.ProductName]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>"); // End inner foreach.
        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a sample XML data source file.
        // -----------------------------------------------------------------
        string xmlContent =
@"<Orders>
    <Order>
        <CustomerName>John Doe</CustomerName>
        <OrderId>1001</OrderId>
        <Items>
            <Item>
                <ProductName>Widget A</ProductName>
                <Quantity>2</Quantity>
            </Item>
            <Item>
                <ProductName>Widget B</ProductName>
                <Quantity>5</Quantity>
            </Item>
        </Items>
    </Order>
    <Order>
        <CustomerName>Jane Smith</CustomerName>
        <OrderId>1002</OrderId>
        <Items>
            <Item>
                <ProductName>Gadget X</ProductName>
                <Quantity>1</Quantity>
            </Item>
        </Items>
    </Order>
</Orders>";
        File.WriteAllText(xmlDataPath, xmlContent);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using the XML data source.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        XmlDataSource xmlDataSource = new XmlDataSource(xmlDataPath);

        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // Default options.

        // Build the report. The data source name ("orders") must match the name used in the template tags.
        engine.BuildReport(loadedTemplate, xmlDataSource, "orders");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}
