using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class JsonReportWithMissingMembers
{
    static void Main()
    {
        // Path to the template document that contains reporting tags.
        string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Path to the JSON file that will be used as the data source.
        string jsonPath = @"C:\Data\ReportData.json";

        // Path where the generated report will be saved.
        string outputPath = @"C:\Reports\GeneratedReport.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create JSON load options (optional – can be omitted if defaults are sufficient).
        JsonDataLoadOptions loadOptions = new JsonDataLoadOptions
        {
            // Example: always generate a root object so that top‑level properties are accessible.
            AlwaysGenerateRootObject = true,
            // Preserve spaces in string values if needed.
            PreserveSpaces = true,
            // Use loose parsing for simple values (numbers, booleans, etc.).
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose
        };

        // Create the JSON data source using the file path and the load options.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath, loadOptions);

        // Initialise the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // Enable the option that treats missing members as null.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: specify a placeholder text for missing members.
            MissingMemberMessage = "N/A"
        };

        // Build the report. The third argument is the name used in the template to reference the data source.
        // If the template does not need to reference the data source object itself, you can pass null or an empty string.
        engine.BuildReport(doc, jsonDataSource, "data");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
