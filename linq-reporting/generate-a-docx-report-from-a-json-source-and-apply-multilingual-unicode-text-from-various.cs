using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings that might be used internally.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "data.json");
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string outputPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample JSON file containing multilingual text fields.
        // -----------------------------------------------------------------
        string jsonContent = @"{
    ""Title"": ""Multilingual Report"",
    ""English"": ""Hello"",
    ""Russian"": ""Привет"",
    ""Chinese"": ""你好"",
    ""Arabic"": ""مرحبا"",
    ""Hindi"": ""नमस्ते""
}";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // ---------------------------------------------------------------
        // 2. Build a Word template programmatically and insert LINQ tags.
        // ---------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title of the report.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Each language line uses a separate tag.
        builder.Writeln("English: <<[model.English]>>");
        builder.Writeln("Russian: <<[model.Russian]>>");
        builder.Writeln("Chinese: <<[model.Chinese]>>");
        builder.Writeln("Arabic: <<[model.Arabic]>>");
        builder.Writeln("Hindi: <<[model.Hindi]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------------------------------------------------------------
        // 3. Load the template and bind the JSON data source.
        // ---------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // JsonDataSource reads the JSON file.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // ReportingEngine processes the template.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // The root object name used in the template tags is "model".
        engine.BuildReport(reportDoc, jsonDataSource, "model");

        // ---------------------------------------------------------------
        // 4. Save the generated report.
        // ---------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
