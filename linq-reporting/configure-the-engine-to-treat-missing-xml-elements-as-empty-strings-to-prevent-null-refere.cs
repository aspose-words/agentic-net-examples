using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class MissingXmlElementsExample
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Person Name: <<[person.Name]>>");
        builder.Writeln("Person Age: <<[person.Age]>>"); // Age element will be missing in the XML.

        // Save the template to a local file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Prepare an XML source where the <Age> element is omitted.
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
    <person>
        <Name>John Doe</Name>
        <!-- Age element is intentionally missing -->
    </person>
</Root>";

        // Load the XML into a memory stream.
        using MemoryStream xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xmlContent));

        // Create an XmlDataSource from the stream.
        XmlDataSource dataSource = new XmlDataSource(xmlStream);

        // Load the template document again (as required by the lifecycle rules).
        Document doc = new Document(templatePath);

        // Configure the ReportingEngine to treat missing members as empty strings.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = string.Empty; // Keep output clean for missing members.

        // Build the report. The root object name is "Root" because the XML root element is <Root>.
        engine.BuildReport(doc, dataSource, "Root");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
