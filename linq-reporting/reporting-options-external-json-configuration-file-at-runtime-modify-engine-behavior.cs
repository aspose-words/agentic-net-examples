using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;

// Define a class that matches the structure of the external JSON configuration.
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

        // Load the reporting configuration from the external JSON file.
        ReportingConfig config = LoadConfig(configPath);

        // Prepare JsonDataLoadOptions based on the configuration.
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions();

        if (config.ExactDateTimeParseFormats != null)
            loadOptions.ExactDateTimeParseFormats = config.ExactDateTimeParseFormats;

        if (config.AlwaysGenerateRootObject.HasValue)
            loadOptions.AlwaysGenerateRootObject = config.AlwaysGenerateRootObject.Value;

        if (config.PreserveSpaces.HasValue)
            loadOptions.PreserveSpaces = config.PreserveSpaces.Value;

        if (!string.IsNullOrEmpty(config.SimpleValueParseMode))
        {
            // Map string to enum value.
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
        if (!string.IsNullOrEmpty(config.ReportBuildOptions))
        {
            // Parse comma‑separated option names into the enum flags.
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
    }

    // Helper method to deserialize the configuration JSON file.
    private static ReportingConfig LoadConfig(string path)
    {
        using FileStream stream = File.OpenRead(path);
        return JsonSerializer.Deserialize<ReportingConfig>(stream);
    }
}
