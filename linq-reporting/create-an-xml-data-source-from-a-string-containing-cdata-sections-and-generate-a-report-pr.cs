using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create an XML file that contains CDATA sections.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<people>
    <person>
        <name><![CDATA[John <Doe>]]></name>
        <bio><![CDATA[Developer & Engineer]]></bio>
    </person>
    <person>
        <name><![CDATA[Jane & Smith]]></name>
        <bio><![CDATA[Designer > Artist]]></bio>
    </person>
</people>";
        string xmlPath = Path.Combine(outputDir, "people.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build a LINQ Reporting template programmatically.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The root name for the data source will be "data".
        // Iterate over each <person> element and output its fields.
        // Since the XML root element is <people>, the generated root object
        // directly exposes the <person> collection, so we iterate over data.person.
        builder.Writeln("<<foreach [p in data.person]>>");
        builder.Writeln("Name: <<[p.name]>>");
        builder.Writeln("Bio: <<[p.bio]>>");
        builder.Writeln("<</foreach>>");

        // Save the template before building the report.
        templateDoc.Save(templatePath);

        // 3. Load the template and bind the XML data source.
        Document reportDoc = new Document(templatePath);

        // Ensure the XML data source generates a root object.
        XmlDataLoadOptions loadOptions = new XmlDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath, loadOptions);

        // 4. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, xmlDataSource, "data");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
