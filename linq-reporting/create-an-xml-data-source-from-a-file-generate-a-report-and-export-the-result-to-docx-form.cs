using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // 1. Create a simple XML data file
        string xmlPath = Path.Combine(outputFolder, "Data.xml");
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
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // 2. Build a template document that contains LINQ Reporting tags
        string templatePath = Path.Combine(outputFolder, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in Person]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 3. Load the template document
        var doc = new Document(templatePath);

        // 4. Load the XML data source with options that always generate a root object
        var loadOptions = new XmlDataLoadOptions { AlwaysGenerateRootObject = true };
        var dataSource = new XmlDataSource(xmlPath, loadOptions);

        // 5. Build the report using the ReportingEngine
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        // The data source name ("people") matches the root element name in the XML.
        engine.BuildReport(doc, dataSource, "people");

        // 6. Save the generated report
        string resultPath = Path.Combine(outputFolder, "Report.docx");
        doc.Save(resultPath);

        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
