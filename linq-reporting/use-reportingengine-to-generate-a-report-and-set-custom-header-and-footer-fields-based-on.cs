using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportData
{
    public string Header { get; set; } = string.Empty;
    public string Footer { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        var json = @"{
            ""Header"": ""Custom Header Text"",
            ""Footer"": ""Custom Footer Text"",
            ""Body"": ""This is the body of the report."",
            ""Items"": [""Item 1"", ""Item 2"", ""Item 3""]
        }";

        // Write JSON to a local file.
        var jsonPath = "data.json";
        File.WriteAllText(jsonPath, json, Encoding.UTF8);

        // Deserialize JSON into the data model.
        var model = JsonConvert.DeserializeObject<ReportData>(File.ReadAllText(jsonPath, Encoding.UTF8))!;

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();

        // Header with custom field.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[model.Header]>>");

        // Footer with custom field.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("<<[model.Footer]>>");

        // Main body content.
        builder.MoveToDocumentEnd();
        builder.Writeln("Report Body:");
        builder.Writeln("<<[model.Body]>>");
        builder.Writeln();
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        doc.Save(outputPath);
    }
}
