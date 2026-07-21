using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple XML data source file.
        const string xmlFileName = "people.xml";
        string xmlContent =
            @"<?xml version=""1.0"" encoding=""utf-8""?>"
          + "<Persons>"
          + "  <Person><Name>John Doe</Name><Age>30</Age></Person>"
          + "  <Person><Name>Jane Smith</Name><Age>25</Age></Person>"
          + "  <Person><Name>Bob Johnson</Name><Age>40</Age></Person>"
          + "</Persons>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        // LINQ Reporting tags: iterate over the collection named "persons".
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Load the XML data source.
        XmlDataSource dataSource = new XmlDataSource(xmlFileName);

        // Build the report. AllowMissingMembers is NOT enabled (default behavior).
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // explicit, but the default is also None.
        engine.BuildReport(template, dataSource, "persons");

        // Save the generated report.
        const string outputFileName = "PeopleReport.docx";
        template.Save(outputFileName);
    }
}
