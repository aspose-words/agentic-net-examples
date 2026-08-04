using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // Reporting engine and JSON data source classes

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data with ISO 8601 date strings.
        string jsonPath = "people.json";
        string jsonContent = @"[
  { ""Name"": ""Alice"", ""BirthDate"": ""1985-03-12T00:00:00"" },
  { ""Name"": ""Bob"",   ""BirthDate"": ""1992-07-25T00:00:00Z"" },
  { ""Name"": ""Carol"", ""BirthDate"": ""2000-11-05T15:30:00+02:00"" }
]";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Create a template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Birth Date: <<[person.BirthDate]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (ensures the engine works on a loaded document).
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template back.
        Document doc = new Document(templatePath);

        // Configure JSON parsing options to recognize ISO 8601 formats explicitly.
        JsonDataLoadOptions jsonOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = new List<string>
            {
                "yyyy-MM-ddTHH:mm:ss",
                "yyyy-MM-ddTHH:mm:ssZ",
                "yyyy-MM-ddTHH:mm:sszzz"
            }
        };

        // Create the JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        bool success = engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
