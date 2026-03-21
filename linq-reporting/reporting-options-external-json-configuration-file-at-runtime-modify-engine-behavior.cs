using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportingConfig
{
    public List<string> ExactDateTimeParseFormats { get; set; }
    public bool? AlwaysGenerateRootObject { get; set; }
    public bool? PreserveSpaces { get; set; }
    public string SimpleValueParseMode { get; set; } // "Loose" or "Strict"
    public string ReportBuildOptions { get; set; }   // e.g., "AllowMissingMembers,RemoveEmptyParagraphs"
}

public class Program
{
    public static void Main()
    {
        // Paths to the files (adjust as needed).
        string templatePath = "template.docx";
        string jsonDataPath = "data.json";
        string configPath   = "config.json";
        string outputPath   = "output.docx";

        // Ensure the template exists; if not, create a minimal one.
        if (!File.Exists(templatePath))
        {
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Hello <<data.Name>>!");
            templateDoc.Save(templatePath);
        }

        // Ensure the JSON data file exists; if not, create a minimal one.
        if (!File.Exists(jsonDataPath))
        {
            var sampleJson = JsonSerializer.Serialize(new { Name = "John Doe" }, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonDataPath, sampleJson);
        }

        // Load the reporting configuration from the external JSON file (if present).
        ReportingConfig config = LoadConfig(configPath);

        // Prepare JsonDataLoadOptions based on the configuration.
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions();

        if (config?.ExactDateTimeParseFormats != null)
            loadOptions.ExactDateTimeParseFormats = config.ExactDateTimeParseFormats;

        if (config?.AlwaysGenerateRootObject.HasValue == true)
            loadOptions.AlwaysGenerateRootObject = config.AlwaysGenerateRootObject.Value;

        if (config?.PreserveSpaces.HasValue == true)
            loadOptions.PreserveSpaces = config.PreserveSpaces.Value;

        if (!string.IsNullOrEmpty(config?.SimpleValueParseMode))
        {
            loadOptions.SimpleValueParseMode = config.SimpleValueParseMode.Equals(
                "Strict", StringComparison.OrdinalIgnoreCase)
                ? JsonSimpleValueParseMode.Strict
                : JsonSimpleValueParseMode.Loose;
        }

        // Create the JSON data source using the configured load options.
        JsonDataSource dataSource = new JsonDataSource(jsonDataPath, loadOptions);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Apply ReportBuildOptions if they are specified in the config.
        if (!string.IsNullOrEmpty(config?.ReportBuildOptions))
        {
            ReportBuildOptions options = ReportBuildOptions.None;
            string[] parts = config.ReportBuildOptions.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                if (Enum.TryParse(part.Trim(), out ReportBuildOptions parsed))
                    options |= parsed;
            }
            engine.Options = options;
        }

        // Build the report using the data source. The data source name "data" can be referenced in the template.
        engine.BuildReport(doc, dataSource, "data");

        // Save the generated report.
        doc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }

    // Helper method to deserialize the configuration JSON file, falling back to defaults if the file is missing or invalid.
    private static ReportingConfig LoadConfig(string path)
    {
        if (!File.Exists(path))
            return new ReportingConfig();

        try
        {
            using FileStream stream = File.OpenRead(path);
            var config = JsonSerializer.Deserialize<ReportingConfig>(stream);
            return config ?? new ReportingConfig();
        }
        catch
        {
            // If deserialization fails, return an empty config to keep the example running.
            return new ReportingConfig();
        }
    }
}
