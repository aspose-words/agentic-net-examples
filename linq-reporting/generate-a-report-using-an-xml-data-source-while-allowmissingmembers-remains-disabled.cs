using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<people>
    <person>
        <name>Alice</name>
        <age>30</age>
    </person>
    <person>
        <name>Bob</name>
        <age>45</age>
    </person>
    <person>
        <name>Charlie</name>
        <age>28</age>
    </person>
</people>";
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "People.xml");
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The XML root (<people>) is treated as a collection of rows.
        // Iterate directly over the root collection.
        builder.Writeln("<<foreach [p in data]>>");
        builder.Writeln("<<[p.name]>> - <<[p.age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Create an XmlDataSource from the XML file.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Initialize the reporting engine without AllowMissingMembers.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build the report. The data source name used in the template is "data".
        engine.BuildReport(reportDoc, xmlDataSource, "data");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
