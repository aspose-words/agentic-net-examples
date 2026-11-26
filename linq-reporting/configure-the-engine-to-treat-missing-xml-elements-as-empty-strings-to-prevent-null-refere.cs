using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML with a missing element (lastName for the first person)
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <person>
        <firstName>John</firstName>
        <!-- lastName element is missing -->
    </person>
    <person>
        <firstName>Jane</firstName>
        <lastName>Doe</lastName>
    </person>
</root>";
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.xml");
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Create a template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a foreach loop over the person nodes
        builder.Writeln("<<foreach [p in data.root.person]>>");
        builder.Writeln("First Name: <<[p.firstName]>>");
        // Missing lastName will be treated as empty string (or inline error message)
        builder.Writeln("Last Name: <<[p.lastName]>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Load XML data source
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Configure the reporting engine to treat missing XML elements as empty strings
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report
        engine.BuildReport(reportDoc, xmlDataSource, "data");

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
