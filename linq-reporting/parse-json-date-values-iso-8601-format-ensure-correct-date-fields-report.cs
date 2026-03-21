using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a minimal Word document to act as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("{{records}}");

        // Sample JSON containing ISO‑8601 date strings.
        string jsonContent = @"[
            { ""date"": ""2023-04-15T13:45:30Z"", ""value"": 123 },
            { ""date"": ""2023-04-16"", ""value"": 456 }
        ]";

        // Write the JSON to a temporary file so Aspose can read it.
        string jsonPath = Path.Combine(Path.GetTempPath(), "Data.json");
        File.WriteAllText(jsonPath, jsonContent);

        // Set up JSON loading options to recognize ISO‑8601 date formats.
        JsonDataLoadOptions jsonOptions = new JsonDataLoadOptions
        {
            ExactDateTimeParseFormats = new List<string>
            {
                "yyyy-MM-ddTHH:mm:ssK", // Full ISO‑8601 with timezone (e.g., Z or +02:00)
                "yyyy-MM-ddTHH:mm:ss",  // Without timezone
                "yyyy-MM-dd"            // Date only
            }
        };

        // Create a JSON data source using the temporary file.
        JsonDataSource dataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "records");

        // Save the generated document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
