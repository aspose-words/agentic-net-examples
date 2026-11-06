using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // XML data source with CDATA sections
        string xmlData = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Report>
  <Title><![CDATA[Report <b>Title</b>]]></Title>
  <Items>
    <Item>
      <Name><![CDATA[Item &amp; One]]></Name>
      <Description><![CDATA[Description with <i>HTML</i> content]]></Description>
    </Item>
    <Item>
      <Name><![CDATA[Second Item]]></Name>
      <Description><![CDATA[Another description]]></Description>
    </Item>
  </Items>
</Report>";

        // Create XML data source from a memory stream (XmlDataSource expects a file path or stream)
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlData));
        XmlDataSource dataSource = new XmlDataSource(xmlStream);

        // Create a template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title (preserve CDATA content as plain text)
        builder.Writeln("<<[report.Title]>>");
        builder.Writeln();

        // Begin foreach over items
        builder.Writeln("<<foreach [item in report.Items.Item]>>");
        builder.Writeln("Name: <<[item.Name]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(template, dataSource, "report");

        // Save the result
        string outputPath = "ReportOutput.docx";
        template.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to {outputPath}");
    }
}
