using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class NestedRepeatingSectionsReport
{
    static void Main()
    {
        // Paths for the template, XML data source and the final report.
        const string templatePath = "OrdersTemplate.docx";
        const string xmlDataPath = "Orders.xml";
        const string reportPath = "OrdersReport.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document that contains nested repeating sections
        //    using the ReportingEngine syntax (<<foreach>> and <<[field]>>).
        // -----------------------------------------------------------------
        Document template = new Document();                     // create a new blank document
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title
        builder.Writeln("Orders Report");
        builder.Writeln();

        // Outer repeating section – iterate over each Order element.
        builder.Writeln("<<foreach [Orders.Order]>>");
        builder.Writeln("Order ID: <<[OrderId]>>");
        builder.Writeln("Customer: <<[Customer]>>");
        builder.Writeln("Items:");

        // Inner repeating section – iterate over each Item inside the current Order.
        builder.Writeln("<<foreach [Items.Item]>>");
        builder.Writeln("- <<[Product]>> x <<[Quantity]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach

        builder.Writeln("<</foreach>>"); // end outer foreach

        // Save the template so it can be loaded later.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare the XML data source.
        //    The XML file contains a list of orders, each with a collection of items.
        // -----------------------------------------------------------------
        const string xmlContent = @"
<Orders>
    <Order>
        <OrderId>1</OrderId>
        <Customer>John Doe</Customer>
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
        <OrderId>2</OrderId>
        <Customer>Jane Smith</Customer>
        <Items>
            <Item>
                <Product>Orange</Product>
                <Quantity>3</Quantity>
            </Item>
        </Items>
    </Order>
</Orders>";

        // Write the XML content to a file (could also use a MemoryStream).
        File.WriteAllText(xmlDataPath, xmlContent);

        // Load the XML data source using the XmlDataSource class.
        XmlDataSource dataSource = new XmlDataSource(xmlDataPath);

        // -----------------------------------------------------------------
        // 3. Build the report by merging the template with the XML data.
        // -----------------------------------------------------------------
        // Load the previously saved template.
        Document reportDoc = new Document(templatePath); // load the template

        // Create a ReportingEngine instance.
        ReportingEngine engine = new ReportingEngine
        {
            // Optional: remove empty paragraphs that may appear after merging.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report. The third parameter is the name used in the template
        // to reference the data source object itself (optional, can be null).
        engine.BuildReport(reportDoc, dataSource, "Orders");

        // -----------------------------------------------------------------
        // 4. Save the final report document.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
