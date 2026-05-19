using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare XML data containing CDATA sections.
        string xmlContent = @"<items>
  <item>
    <title><![CDATA[Hello <World>]]></title>
    <description><![CDATA[This is a description with special characters: & < >]]></description>
  </item>
  <item>
    <title><![CDATA[Second Item]]></title>
    <description><![CDATA[Another description]]></description>
  </item>
</items>";

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from XML with CDATA sections:");
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Title: <<[item.title]>>");
        builder.Writeln("Description: <<[item.description]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Create an XmlDataSource from the XML string.
        using (MemoryStream xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent)))
        {
            XmlDataSource dataSource = new XmlDataSource(xmlStream);

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(reportDoc, dataSource, "items");
        }

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
