using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Header.
        builder.Writeln("Customer Report");
        builder.Writeln("----------------");

        // Use a foreach loop over a collection named "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        // The XML may miss the Age element; we want it to appear as an empty string.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age:  <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        template.Save(templatePath);

        // Load the template back (simulating a real scenario where the template is a file).
        Document loadedTemplate = new Document(templatePath);

        // Sample XML data: the second person lacks the <Age> element.
        string xmlContent = @"
<persons>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <!-- Age element is missing for Jane -->
    </person>
</persons>";

        // Create an XmlDataSource from the XML string.
        using MemoryStream xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xmlContent));
        XmlDataSource xmlDataSource = new XmlDataSource(xmlStream);

        // Configure the ReportingEngine to treat missing members as empty strings.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        // Optional: customize the message shown for missing members (empty string suppresses any output).
        engine.MissingMemberMessage = string.Empty;

        // Build the report. The data source name must match the collection name used in the template.
        engine.BuildReport(loadedTemplate, xmlDataSource, "persons");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "report.docx");
        loadedTemplate.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
