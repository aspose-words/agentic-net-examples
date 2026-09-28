using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample JSON containing a root array.
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""Bob"",   ""Age"": 25 },
            { ""Name"": ""Carol"", ""Age"": 28 }
        ]";

        // Write JSON to a temporary file so JsonDataSource can read it.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        File.WriteAllText(jsonPath, json);

        // Create the JSON data source. In this version of Aspose.Words the
        // JsonDataLoadOptions class does not expose a FlattenRootArray property,
        // so we use the default options which already allow iterating over the root array.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // LINQ Reporting tags: iterate over the root array and output fields.
        builder.Writeln("<<foreach [person in data]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Create and configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report using the JSON data source.
        engine.BuildReport(report, dataSource, "data");

        // Save the generated report.
        string reportPath = "Report.docx";
        report.Save(reportPath);
    }
}
