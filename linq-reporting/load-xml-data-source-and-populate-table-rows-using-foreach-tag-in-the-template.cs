using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data source file.
        const string xmlPath = "people.xml";
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<People>
    <Person>
        <Name>John Doe</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Jane Smith</Name>
        <Age>25</Age>
    </Person>
    <Person>
        <Name>Bob Johnson</Name>
        <Age>40</Age>
    </Person>
</People>";
        File.WriteAllText(xmlPath, xmlContent);

        // Build the template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People List:");
        // The data source name is "persons", so the foreach must iterate over that collection.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        template.Save(templatePath);

        // Load the template for report generation.
        Document report = new Document(templatePath);

        // Create an XML data source from the file.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the data source. The root name "persons" must match the name used in BuildReport.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(report, dataSource, "persons");

        // Save the generated report.
        const string outputPath = "output.docx";
        report.Save(outputPath);
    }
}
