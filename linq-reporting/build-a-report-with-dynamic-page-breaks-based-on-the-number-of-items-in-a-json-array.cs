using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        string jsonContent = @"[
            { ""Name"": ""Item 1"" },
            { ""Name"": ""Item 2"" },
            { ""Name"": ""Item 3"" },
            { ""Name"": ""Item 4"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Iterate over each item in the JSON array and insert a page break after each item.
        builder.Writeln("<<foreach [item in data]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Load JSON data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
