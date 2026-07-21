using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // File paths.
        const string templatePath = "Template.docx";
        const string jsonPath = "Data.json";
        const string outputPath = "Report.docx";

        // Sample JSON data (array of objects with identical keys).
        string json = @"[
            { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""New York"" },
            { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""Los Angeles"" },
            { ""Name"": ""Carol"", ""Age"": 28, ""City"": ""Chicago"" }
        ]";
        File.WriteAllText(jsonPath, json);

        // -----------------------------------------------------------------
        // 1. Build the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report generated from JSON data:");
        builder.Writeln(); // Empty paragraph.

        // Begin the foreach block.
        builder.Writeln("<<foreach [item in data]>>");

        // Start a table.
        Table table = builder.StartTable();

        // Header row – literal strings via expression tags.
        builder.InsertCell();
        builder.Writeln("<<[\"Name\"]>>");
        builder.InsertCell();
        builder.Writeln("<<[\"Age\"]>>");
        builder.InsertCell();
        builder.Writeln("<<[\"City\"]>>");
        builder.EndRow();

        // Data row – repeat for each JSON object.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.City]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report using the JSON data source.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var jsonDataSource = new JsonDataSource(jsonPath);

        var engine = new ReportingEngine();
        // The root object name used in the template tags is "data".
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
