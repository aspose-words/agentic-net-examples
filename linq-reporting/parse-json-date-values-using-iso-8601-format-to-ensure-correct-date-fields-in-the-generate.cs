using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data with ISO 8601 date strings.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "people.json");
        File.WriteAllText(jsonPath,
            @"[
                { ""Name"": ""John Doe"", ""BirthDate"": ""1990-05-15T00:00:00Z"" },
                { ""Name"": ""Jane Smith"", ""BirthDate"": ""1985-12-01T00:00:00Z"" }
            ]");

        // Create a template document containing LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Birth Date: <<[p.BirthDate]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure JSON parsing to recognize ISO 8601 date formats.
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = new List<string>
            {
                "yyyy-MM-ddTHH:mm:ssZ",
                "yyyy-MM-ddTHH:mm:sszzz",
                "yyyy-MM-ddTHH:mm:ss"
            }
        };

        // Create a JSON data source using the options above.
        JsonDataSource dataSource = new JsonDataSource(jsonPath, loadOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, dataSource, "persons");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
