using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(outputDir);

        // Create a sample JSON file with some missing members.
        string jsonPath = Path.Combine(dataDir, "people.json");
        string jsonContent = @"{
  ""persons"": [
    { ""Name"": ""John Doe"", ""Age"": 30 },
    { ""Name"": ""Jane Smith"" },
    { ""Name"": ""Bob Johnson"", ""Age"": 45, ""City"": ""New York"" }
  ]
}";
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // Build a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a foreach block that iterates over the "persons" collection.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.City]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "N/A"
        };

        // Load the JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report. The root object name in the template is "persons".
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "ReportFromJson.docx");
        doc.Save(outputPath);
    }
}
