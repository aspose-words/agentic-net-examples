using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title: <<[report.Title]>>");
        builder.Writeln("Report Content: <<[report.Content]>>");
        templateDoc.Save(templatePath);

        // 2. Prepare XML data containing CDATA sections.
        var xmlContent = @"<?xml version='1.0' encoding='utf-8'?>
<Report>
  <Title><![CDATA[Sample <Report> Title]]></Title>
  <Content><![CDATA[This is the content with special characters: <, >, &, ""quotes"".]]></Content>
</Report>";
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(xmlContent));
        xmlStream.Position = 0; // Ensure the stream is ready for reading.

        // 3. Load the template document.
        var doc = new Document(templatePath);

        // 4. Create an XmlDataSource from the XML stream.
        var xmlDataSource = new XmlDataSource(xmlStream);

        // 5. Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, xmlDataSource, "report");

        // 6. Save the generated report.
        var outputPath = "ReportOutput.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input required).
        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
