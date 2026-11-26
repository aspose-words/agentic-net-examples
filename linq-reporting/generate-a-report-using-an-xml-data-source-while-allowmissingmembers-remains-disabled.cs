using System;
using System.Data;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingXmlExample
{
    public static void Main()
    {
        // Register code page provider for XML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample XML data.
        // -----------------------------------------------------------------
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath,
            @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Orders>
  <Order>
    <CustomerName>John Doe</CustomerName>
    <OrderDate>2023-01-01</OrderDate>
    <Items>
      <Item>
        <ProductName>Widget</ProductName>
        <Quantity>2</Quantity>
      </Item>
      <Item>
        <ProductName>Gadget</ProductName>
        <Quantity>5</Quantity>
      </Item>
    </Items>
  </Order>
  <Order>
    <CustomerName>Jane Smith</CustomerName>
    <OrderDate>2023-02-15</OrderDate>
    <Items>
      <Item>
        <ProductName>Thingamajig</ProductName>
        <Quantity>1</Quantity>
      </Item>
    </Items>
  </Order>
</Orders>");

        // -----------------------------------------------------------------
        // 2. Build a template document that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Orders Report");
        builder.Writeln("<<foreach [order in Order]>>");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.ProductName]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template document for reporting.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 4. Load XML data into a DataSet (the format expected by the engine).
        // -----------------------------------------------------------------
        DataSet dataSet = new DataSet();
        dataSet.ReadXml(xmlPath);

        // -----------------------------------------------------------------
        // 5. Build the report. AllowMissingMembers stays disabled (default).
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // No special options are set; the default disables AllowMissingMembers.
        engine.BuildReport(doc, dataSet, "Orders");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
