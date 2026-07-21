using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for the Table class

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        // Paths for JSON data, template, and final report.
        string jsonPath = Path.Combine(outputDir, "data.json");
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // Create sample JSON data.
        string json = @"[
  { ""Index"": 1, ""Name"": ""Alice"" },
  { ""Index"": 2, ""Name"": ""Bob"" },
  { ""Index"": 3, ""Name"": ""Charlie"" }
]";
        File.WriteAllText(jsonPath, json);

        // Build the template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Optional title.
        builder.Writeln("Report generated from JSON");

        // Table header (static, appears once).
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();
        builder.EndTable();

        // Begin foreach block over the JSON array named "items".
        builder.Writeln("<<foreach [item in items]>>");

        // Table row for each item.
        Table dataTable = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Create a JSON data source pointing to the file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source. The root name "items" matches the foreach tag.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, jsonDataSource, "items");

        // Save the final document.
        reportDoc.Save(resultPath);
    }
}
