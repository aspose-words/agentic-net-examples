using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Create sample XML with a namespace and save it.
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
<root xmlns:ns=""http://example.com/ns"">
    <ns:Person>
        <ns:Name>John Doe</ns:Name>
        <ns:Age>30</ns:Age>
    </ns:Person>
    <ns:Person>
        <ns:Name>Jane Smith</ns:Name>
        <ns:Age>25</ns:Age>
    </ns:Person>
</root>";
        string xmlPath = Path.Combine(workDir, "data.xml");
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Build a template document that uses LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        builder.Writeln();

        // Use LINQ Reporting tags without the namespace prefix.
        // The XmlDataLoadOptions will generate a root object so that the elements can be accessed directly.
        builder.Writeln("<<foreach [person in root.Person]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(workDir, "template.docx");
        template.Save(templatePath);

        // 3. Load the template (simulating a separate step).
        Document loadedTemplate = new Document(templatePath);

        // 4. Create an XmlDataSource from the XML file with options that always generate a root object.
        XmlDataLoadOptions loadOptions = new XmlDataLoadOptions
        {
            AlwaysGenerateRootObject = true
        };
        XmlDataSource xmlDataSource = new XmlDataSource(xmlPath, loadOptions);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine();
        // The root object name in the template is "root".
        engine.BuildReport(loadedTemplate, xmlDataSource, "root");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "report.docx");
        loadedTemplate.Save(reportPath);
    }
}
