using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Iterate over each <person> element.
        builder.Writeln("<<foreach [p in person]>>");
        // This element exists in the XML.
        builder.Writeln("Age: <<[p.Age]>>");
        // This element is missing in the XML; it will be treated as an empty string.
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Create an XML data source where the <Name> element is omitted.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<people>
    <person>
        <Age>45</Age>
    </person>
    <person>
        <Age>30</Age>
        <Name>John Doe</Name>
    </person>
</people>";
        const string xmlPath = "Data.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Configure the reporting engine to treat missing members as empty strings.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Build the report. The root object name matches the top‑level XML element.
        engine.BuildReport(doc, dataSource, "people");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
