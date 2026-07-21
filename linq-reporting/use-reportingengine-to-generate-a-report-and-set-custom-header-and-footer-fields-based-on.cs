using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for temporary files.
        string jsonPath = "reportData.json";
        string templatePath = "ReportTemplate.docx";
        string outputPath = "ReportResult.docx";

        // 1. Create sample JSON data.
        var sampleData = new ReportModel
        {
            Header = "Custom Header Text",
            Footer = "Custom Footer Text",
            Body = "This is the main body of the report generated from JSON data."
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented), Encoding.UTF8);

        // 2. Build the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Body content placeholder.
        builder.Writeln("<<[model.Body]>>");

        // Header placeholder.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[model.Header]>>");

        // Footer placeholder.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("<<[model.Footer]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // 4. Create a JSON data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using ReportingEngine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(reportDoc, jsonDataSource, "model");

        // 6. Save the generated report.
        reportDoc.Save(outputPath);
    }
}

// Data model that matches the JSON structure.
public class ReportModel
{
    public string Header { get; set; } = string.Empty;
    public string Footer { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
}
