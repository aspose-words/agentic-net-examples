using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string jsonPath = Path.Combine(outputDir, "data.json");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // 1. Create sample JSON data.
        string jsonContent = @"{
  ""Header"": ""Custom Header Text"",
  ""Footer"": ""Custom Footer Text"",
  ""Title"": ""LINQ Reporting Example"",
  ""Body"": ""This report demonstrates how to set header and footer fields from JSON data using Aspose.Words ReportingEngine.""
}";
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Build the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header tag.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("<<[model.Header]>>");

        // Footer tag.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("<<[model.Footer]>>");

        // Body of the document.
        builder.MoveToSection(0);
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();
        builder.Writeln("<<[model.Body]>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // 4. Create a JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(reportDoc, jsonDataSource, "model");

        // 6. Save the generated report.
        reportDoc.Save(reportPath);
    }
}
