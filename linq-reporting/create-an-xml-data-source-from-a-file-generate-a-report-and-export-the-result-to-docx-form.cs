using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML handling.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string workingDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workingDir, "Template.docx");
        string xmlPath = Path.Combine(workingDir, "Data.xml");
        string outputPath = Path.Combine(workingDir, "Report.docx");

        // Create a sample XML data file.
        CreateSampleXml(xmlPath);

        // Create a Word template containing LINQ Reporting tags.
        CreateTemplateDocument(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlPath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Save the generated report.
        doc.Save(outputPath);
    }

    private static void CreateSampleXml(string filePath)
    {
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<People>
  <Person>
    <Name>John Doe</Name>
    <Age>30</Age>
  </Person>
  <Person>
    <Name>Jane Smith</Name>
    <Age>25</Age>
  </Person>
</People>";
        File.WriteAllText(filePath, xmlContent);
    }

    private static void CreateTemplateDocument(string filePath)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Title.
        builder.Writeln("People Report");
        builder.Writeln();

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        template.Save(filePath);
    }
}
