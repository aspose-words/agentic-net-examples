using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // File paths.
        const string jsonPath = "people.json";
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // Create a sample JSON file.
        string jsonContent = @"[
  { ""Name"": ""Alice"", ""Age"": 30, ""BirthDate"": ""1992-05-15T00:00:00"" },
  { ""Name"": ""Bob"", ""Age"": 45, ""BirthDate"": ""1977-11-23T00:00:00"" }
]";
        File.WriteAllText(jsonPath, jsonContent);

        // Build a template document containing LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated from JSON data:");
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        // Removed the invalid ':d' format switch.
        builder.Writeln("Birth date: <<[person.BirthDate]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (required before loading for the report).
        templateDoc.Save(templatePath);

        // Load the template document.
        Document loadedTemplate = new Document(templatePath);

        // Configure JSON data load options (type inference is enabled by default).
        JsonDataLoadOptions jsonOptions = new JsonDataLoadOptions
        {
            PreserveSpaces = true,
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose
        };

        // Create a JsonDataSource from the file using the options.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, jsonDataSource, "persons");

        // Save the generated report.
        loadedTemplate.Save(outputPath);
    }
}
