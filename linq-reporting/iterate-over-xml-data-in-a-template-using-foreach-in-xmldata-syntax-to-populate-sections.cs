using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFileName = "people.xml";
        string xmlContent =
            @"<People>" +
            @"  <Person>" +
            @"    <Name>John Doe</Name>" +
            @"    <Age>30</Age>" +
            @"  </Person>" +
            @"  <Person>" +
            @"    <Name>Jane Smith</Name>" +
            @"    <Age>25</Age>" +
            @"  </Person>" +
            @"</People>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Create a template document with LINQ Reporting tags.
        const string templateFileName = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People List:");
        // Iterate over the Person elements from the XML data source named "xmlData".
        // When the root element contains a collection, the data source itself can be iterated directly.
        builder.Writeln("<<foreach [person in xmlData]>>");
        builder.Writeln("- Name: <<[person.Name]>>");
        builder.Writeln("- Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templateFileName);

        // Load the template for report generation.
        Document reportDoc = new Document(templateFileName);

        // Create an XML data source.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlFileName);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options required.
        bool success = engine.BuildReport(reportDoc, xmlDataSource, "xmlData");

        // Save the generated report.
        const string outputFileName = "output.docx";
        reportDoc.Save(outputFileName);

        // Indicate completion.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
