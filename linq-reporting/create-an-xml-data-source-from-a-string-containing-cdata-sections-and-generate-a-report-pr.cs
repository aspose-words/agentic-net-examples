using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for UTF‑8 support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // XML string that contains CDATA sections.
        string xml = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<root>
    <Item>
        <Content><![CDATA[Some <b>bold</b> text]]></Content>
    </Item>
    <Item>
        <Content><![CDATA[Another & special <i>italic</i> text]]></Content>
    </Item>
</root>";

        // Load the XML into a memory stream and create an XmlDataSource.
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        var xmlDataSource = new XmlDataSource(xmlStream);

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from XML with CDATA values:");
        // Iterate over the rows that represent <Item> elements.
        builder.Writeln("<<foreach [item in root]>>");
        builder.Writeln("Content: <<[item.Content]>>");
        builder.Writeln("<</foreach>>");

        // Save the template so it can be loaded for reporting.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // BuildReport uses the XML data source; the root name is "root".
        engine.BuildReport(reportDoc, xmlDataSource, "root");

        // Save the final report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
