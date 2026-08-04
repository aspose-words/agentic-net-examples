using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // XML string with CDATA sections.
        string xmlContent = @"<?xml version='1.0' encoding='utf-8'?>
<Root>
    <Item>
        <Description><![CDATA[Some <b>bold</b> text]]></Description>
    </Item>
    <Item>
        <Description><![CDATA[Another <i>italic</i> text]]></Description>
    </Item>
</Root>";

        // Load XML into a memory stream.
        using MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
        xmlStream.Position = 0;

        // Ensure the root object is always generated for proper collection access.
        XmlDataLoadOptions loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
        XmlDataSource dataSource = new XmlDataSource(xmlStream, loadOptions);

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Items Report");
        // LINQ Reporting foreach loop over the collection Root.Item
        builder.Writeln("<<foreach [item in root.Item]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the XML data source; the root object name is "root".
        engine.BuildReport(report, dataSource, "root");

        // Save the generated report.
        const string reportPath = "Report.docx";
        report.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
