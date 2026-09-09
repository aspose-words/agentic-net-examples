using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // File names.
        string templatePath = "Template.docx";
        string xmlPath = "People.xml";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple XML data source file.
        // -----------------------------------------------------------------
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<people>
    <person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </person>
    <person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </person>
</people>";
        File.WriteAllText(xmlPath, xmlContent);

        // -----------------------------------------------------------------
        // 2. Build a template document that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Iterate over the collection of <person> elements.
        builder.Writeln("<<foreach [p in people]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and generate the report using the XML data.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create an XmlDataSource from the XML file.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath);

        // Initialise the reporting engine without AllowMissingMembers.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default behavior; missing members cause an exception.

        // Build the report. The data source name must match the root element name used in the template.
        engine.BuildReport(reportDoc, xmlDataSource, "people");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
