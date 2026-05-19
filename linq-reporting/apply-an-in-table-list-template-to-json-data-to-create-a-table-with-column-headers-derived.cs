using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ---------- Step 1: Create sample JSON data ----------
        string json = @"
[
    { ""Name"": ""Alice"", ""Age"": 30, ""City"": ""New York"" },
    { ""Name"": ""Bob"",   ""Age"": 25, ""City"": ""London"" },
    { ""Name"": ""Charlie"", ""Age"": 28, ""City"": ""Paris"" }
]";
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        File.WriteAllText(jsonPath, json);

        // ---------- Step 2: Build the template document ----------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin the foreach block that iterates over the JSON array.
        builder.Writeln("<<foreach [item in data]>>");

        // Create a table. Header row is static (derived from the JSON keys).
        Table table = builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("City");
        builder.EndRow();

        // Data row – values are filled from the JSON objects.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.City]>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------- Step 3: Load the template and build the report ----------
        Document reportDoc = new Document(templatePath);

        // Create a JSON data source from the file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Configure and run the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // explicit assignment as required.
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // ---------- Step 4: Save the generated report ----------
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
