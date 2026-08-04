using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple template with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Person Report");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>"); // Age may be missing in the XML.
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Create sample XML data where the first person lacks the <Age> element.
        var xmlPath = "Data.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<persons>
    <person>
        <Name>John Doe</Name>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>30</Age>
    </person>
</persons>";
        File.WriteAllText(xmlPath, xmlContent);

        // Load the template document.
        var doc = new Document(templatePath);

        // Load the XML data source.
        var dataSource = new XmlDataSource(xmlPath);

        // Configure the reporting engine to treat missing members as empty strings.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        // Optional: customize the message shown for missing members (empty string suppresses output).
        engine.MissingMemberMessage = string.Empty;

        // Build the report. The data source name must match the collection used in the template.
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }
}
