using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = "items.json";
        string jsonContent = @"[
            { ""Name"": ""Item 1"" },
            { ""Name"": ""Item 2"" },
            { ""Name"": ""Item 3"" },
            { ""Name"": ""Item 4"" },
            { ""Name"": ""Item 5"" }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin a foreach block that iterates over the JSON array (named "items").
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Item: <<[item.Name]>>");
        // Insert a page break after each item – this will be repeated for every iteration.
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back for reporting.
        var doc = new Document(templatePath);

        // Create a JSON data source from the file.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        // The data source name must match the name used in the template ("items").
        engine.BuildReport(doc, jsonDataSource, "items");

        // Save the generated report.
        doc.Save("ReportWithPageBreaks.docx");
    }
}
