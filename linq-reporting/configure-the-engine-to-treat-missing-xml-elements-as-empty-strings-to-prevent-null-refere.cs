using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a simple template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // The template contains a tag that references a missing XML element (Name).
        // When the XML data does not contain this element, we want the engine to treat it as an empty string.
        builder.Writeln("Customer information:");
        builder.Writeln("Name: <<[person.Name]>>");   // Missing element.
        builder.Writeln("Age: <<[person.Age]>>");     // Present element.

        // Save the template to a temporary file (required by the lifecycle rules).
        string templatePath = Path.Combine(Path.GetTempPath(), "Template.docx");
        template.Save(templatePath);

        // Prepare XML data that lacks the <Name> element.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<root>
    <person>
        <Age>30</Age>
    </person>
</root>";
        // Write XML to a temporary file.
        string xmlPath = Path.Combine(Path.GetTempPath(), "Data.xml");
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create an XmlDataSource from the XML file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Configure the ReportingEngine to treat missing members as empty strings.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Optional: customize the message shown for missing members (empty string means nothing is inserted).
        engine.MissingMemberMessage = string.Empty;

        // Build the report. The root object name is empty because we are using a data source directly.
        engine.BuildReport(doc, dataSource, "root");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);

        // Inform the user (no interactive input required).
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
