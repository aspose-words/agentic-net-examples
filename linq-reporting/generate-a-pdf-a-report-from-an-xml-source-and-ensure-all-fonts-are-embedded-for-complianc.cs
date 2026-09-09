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
        // Register code page provider for XML parsing (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        const string xmlFileName = "order.xml";
        File.WriteAllText(xmlFileName,
@"<order>
    <CustomerName>John Doe</CustomerName>
    <Items>
        <Item>
            <Name>Product A</Name>
            <Price>10.5</Price>
        </Item>
        <Item>
            <Name>Product B</Name>
            <Price>20</Price>
        </Item>
    </Items>
</order>");

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Customer: <<[order.CustomerName]>>");
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in order.Items.Item]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        var xmlDataSource = new XmlDataSource(xmlFileName);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, xmlDataSource, "order");

        // Configure PDF/A save options with full font embedding.
        var pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true,
            FontEmbeddingMode = PdfFontEmbeddingMode.EmbedAll,
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the generated report as PDF/A.
        const string outputPdf = "Report.pdf";
        template.Save(outputPdf, pdfOptions);
    }
}
