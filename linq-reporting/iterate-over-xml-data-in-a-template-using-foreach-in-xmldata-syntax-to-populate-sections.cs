using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Persons>
    <Person>
        <Name>Alice</Name>
        <Age>30</Age>
    </Person>
    <Person>
        <Name>Bob</Name>
        <Age>25</Age>
    </Person>
    <Person>
        <Name>Charlie</Name>
        <Age>35</Age>
    </Person>
</Persons>";
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML data source.
        XmlDataSource xmlData = new XmlDataSource(xmlPath);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        // Iterate over the rows of the Persons table.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the XML data source.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(template, xmlData, "Persons");

        // Save the generated report.
        string outputPath = "report.docx";
        template.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
