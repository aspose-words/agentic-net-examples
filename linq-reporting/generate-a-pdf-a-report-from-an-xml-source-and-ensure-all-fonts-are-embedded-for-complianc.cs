using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "order.xml");
        File.WriteAllText(xmlPath, @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Order>
    <CustomerName>John Doe</CustomerName>
    <OrderDate>2023-08-15</OrderDate>
    <Items>
        <Item>
            <Name>Widget A</Name>
            <Quantity>2</Quantity>
            <Price>19.99</Price>
        </Item>
        <Item>
            <Name>Gadget B</Name>
            <Quantity>1</Quantity>
            <Price>99.50</Price>
        </Item>
        <Item>
            <Name>Thingamajig C</Name>
            <Quantity>5</Quantity>
            <Price>5.75</Price>
        </Item>
    </Items>
</Order>");

        // 2. Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Date: <<[order.OrderDate]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>> | Qty: <<[item.Quantity]>> | Price: $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Calculate total using a simple expression.
        builder.Writeln("Total: $<<[order.Items.Item.Sum(i => i.Quantity * i.Price)]>>");

        doc.Save(templatePath);

        // 3. Load the template.
        var reportDoc = new Document(templatePath);

        // 4. Prepare the XML data source.
        var xmlDataSource = new XmlDataSource(xmlPath);

        // 5. Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, xmlDataSource, "order");

        // 6. Save as PDF/A with all fonts embedded.
        string pdfPath = Path.Combine(outputDir, "OrderReport.pdf");
        var pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };
        reportDoc.Save(pdfPath, pdfOptions);
    }
}
