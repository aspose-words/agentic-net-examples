using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable code pages for Unicode handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON containing multilingual text.
        string jsonContent = @"{
  ""Title"": ""Multilingual Report"",
  ""Description"": ""Demonstration of LINQ Reporting with Unicode text."",
  ""Languages"": [
    { ""LanguageName"": ""English"", ""Text"": ""Hello, world!"" },
    { ""LanguageName"": ""Русский"", ""Text"": ""Привет, мир!"" },
    { ""LanguageName"": ""中文"", ""Text"": ""你好，世界！"" },
    { ""LanguageName"": ""العربية"", ""Text"": ""مرحبا بالعالم!"" },
    { ""LanguageName"": ""हिन्दी"", ""Text"": ""नमस्ते दुनिया!"" }
  ]
}";
        // Write JSON to a local file.
        string dataFile = "data.json";
        File.WriteAllText(dataFile, jsonContent, Encoding.UTF8);

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<<[model.Description]>>");
        builder.Writeln();

        builder.Writeln("Languages:");
        builder.Writeln("<<foreach [lang in Languages]>>");
        builder.Writeln("- <<[lang.LanguageName]>>: <<[lang.Text]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templateFile = "template.docx";
        template.Save(templateFile);

        // Load the template for reporting.
        Document reportDoc = new Document(templateFile);

        // Load JSON data source.
        Aspose.Words.Reporting.JsonDataSource jsonData = new Aspose.Words.Reporting.JsonDataSource(dataFile);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(reportDoc, jsonData, "model");

        // Save the generated report.
        string outputFile = "report.docx";
        reportDoc.Save(outputFile);
    }
}
