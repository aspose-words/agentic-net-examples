using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class LinqReportingExample
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFileName = "People.xml";
        const string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
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
        File.WriteAllText(xmlFileName, xmlContent);

        // Create a template document with LINQ Reporting tags.
        const string templateFileName = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("----------------");
        // Begin foreach loop over the XML data source named "xmlData".
        builder.Writeln("<<foreach [in xmlData]>>");
        // Inside the loop output the fields of each Person element.
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Age: <<[Age]>>");
        builder.Writeln(""); // Blank line between records.
        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templateFileName);

        // Load the template document for reporting.
        Document reportDoc = new Document(templateFileName);

        // Create an XML data source from the XML file.
        XmlDataSource xmlDataSource = new XmlDataSource(xmlFileName);

        // Build the report using the data source. The data source name must match the tag ("xmlData").
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, xmlDataSource, "xmlData");

        // Save the generated report.
        const string outputFileName = "Report.docx";
        reportDoc.Save(outputFileName);

        Console.WriteLine(success
            ? $"Report generated successfully: {outputFileName}"
            : "Report generation failed.");
    }
}
