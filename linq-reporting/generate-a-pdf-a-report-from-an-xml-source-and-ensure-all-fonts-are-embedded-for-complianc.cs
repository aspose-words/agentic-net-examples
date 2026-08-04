using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
    <CustomerName>John Doe</CustomerName>
    <OrderDate>2023-01-15</OrderDate>
    <Items>
        <Item>
            <Name>Apple</Name>
            <Quantity>5</Quantity>
        </Item>
        <Item>
            <Name>Banana</Name>
            <Quantity>3</Quantity>
        </Item>
        <Item>
            <Name>Orange</Name>
            <Quantity>7</Quantity>
        </Item>
    </Items>
</Report>";
        string xmlPath = "ReportData.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Create a blank Word document and insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Customer: <<[data.CustomerName]>>");
        builder.Writeln("Order Date: <<[data.OrderDate]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in data.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Quantity]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "data");

        // Configure PDF/A save options with full font embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            EmbedFullFonts = true,
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll
        };

        // Save the final document as PDF/A.
        doc.Save("Report.pdf", saveOptions);
    }
}
