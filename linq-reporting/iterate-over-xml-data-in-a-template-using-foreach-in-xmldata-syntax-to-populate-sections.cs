using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        string dataPath = "people.xml";
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // 1. Create sample XML data.
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>
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
        File.WriteAllText(dataPath, xmlContent);

        // 2. Build a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("----------------");
        // Use the foreach syntax that iterates over the XML data source.
        builder.Writeln("<<foreach [in xmlData]>>");
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Age: <<[Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("----------------");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template back (simulating a separate load step).
        Document loadedTemplate = new Document(templatePath);

        // 4. Create an XmlDataSource from the XML file.
        XmlDataSource xmlDataSource = new XmlDataSource(dataPath);

        // 5. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source name used in the template tags is "xmlData".
        engine.BuildReport(loadedTemplate, xmlDataSource, "xmlData");

        // 6. Save the generated report.
        loadedTemplate.Save(outputPath);

        // Optional: indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
