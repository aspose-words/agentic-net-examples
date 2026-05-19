using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory for temporary files.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample XML without a default namespace.
        string xmlContent = @"<people>
    <person>
        <name>Alice</name>
        <age>30</age>
    </person>
    <person>
        <name>Bob</name>
        <age>45</age>
    </person>
</people>";
        string xmlPath = Path.Combine(workDir, "people.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // The root data source name is "data". Iterate over each <person> element.
        builder.Writeln("<<foreach [p in data.person]>>");
        builder.Writeln("Name: <<[p.name]>>");
        builder.Writeln("Age: <<[p.age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (demonstrates the required save step).
        string templatePath = Path.Combine(workDir, "Template.docx");
        template.Save(templatePath);

        // 3. Load the template (required before building the report).
        Document reportDoc = new Document(templatePath);

        // 4. Create an XmlDataSource with options that keep the root element.
        XmlDataLoadOptions loadOptions = new XmlDataLoadOptions
        {
            // Ensure the root element ("people") is exposed as an object.
            AlwaysGenerateRootObject = true
        };
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath, loadOptions);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "data".
        engine.BuildReport(reportDoc, xmlDataSource, "data");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workDir, "PeopleReport.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine("Report generated: " + outputPath);
    }
}
