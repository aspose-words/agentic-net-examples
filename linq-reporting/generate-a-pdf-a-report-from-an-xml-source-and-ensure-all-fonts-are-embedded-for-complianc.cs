using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlContent =
            @"<order>
                <CustomerName>John Doe</CustomerName>
                <Items>
                    <Item>
                        <Name>Widget A</Name>
                        <Price>19.99</Price>
                    </Item>
                    <Item>
                        <Name>Gadget B</Name>
                        <Price>29.49</Price>
                    </Item>
                </Items>
              </order>";
        // Write XML to a temporary file.
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "order.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // Create a blank Word document and build the LINQ Reporting template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Order Report");
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        using (FileStream xmlStream = File.OpenRead(xmlPath))
        {
            XmlDataSource dataSource = new XmlDataSource(xmlStream);
            // Build the report using the data source; the root object name is "order".
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "order");
        }

        // Configure PDF/A save options with full font embedding.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b,
            EmbedFullFonts = true
        };

        // Save the document as PDF/A.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OrderReport.pdf");
        doc.Save(outputPath, saveOptions);
    }
}
