using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Enable reflection optimization for the ReportingEngine.
        ReportingEngine.UseReflectionOptimization = true;

        // Ensure a template document exists.
        const string templatePath = "Template.docx";
        if (!File.Exists(templatePath))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Report generated on {{DateTime.Now}}");
            doc.Save(templatePath);
        }

        // Ensure a simple XML data source exists.
        const string xmlPath = "LargeData.xml";
        if (!File.Exists(xmlPath))
        {
            const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <Item>SampleValue</Item>
</root>";
            File.WriteAllText(xmlPath, xmlContent);
        }

        // Create a new instance of the ReportingEngine.
        var engine = new ReportingEngine();

        // Register a known external type (e.g., System.DateTime) so its static members can be accessed
        // directly from the report template.
        engine.KnownTypes.Add(typeof(DateTime));

        // Load the template document that contains the report placeholders.
        var template = new Document(templatePath);

        // Load the XML data set.
        var xmlData = new XmlDataSource(xmlPath);

        // Build the report. The third argument is the name of the top‑level XML element that
        // serves as the root for the data source.
        engine.BuildReport(template, xmlData, "root");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        template.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
