using System;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;

class ReflectionOptimizationDemo
{
    // Threshold for "small" JSON arrays. Adjust as needed.
    private const int SmallArrayThreshold = 10;

    static void Main()
    {
        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Create temporary files for the demo.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeDemo");
        Directory.CreateDirectory(tempDir);

        // Path to the template document.
        string templatePath = Path.Combine(tempDir, "Template.docx");

        // Create a minimal template document.
        Document templateDoc = new Document();
        templateDoc.Save(templatePath);

        // Path to the JSON data file.
        string jsonPath = Path.Combine(tempDir, "People.json");

        // Write sample JSON data.
        string sampleJson = @"{
            ""persons"": [
                { ""Name"": ""John"", ""Age"": 30 },
                { ""Name"": ""Jane"", ""Age"": 25 }
            ]
        }";
        File.WriteAllText(jsonPath, sampleJson);

        // Load the template.
        Document doc = new Document(templatePath);

        // Determine whether the JSON array is small.
        bool isSmallArray = IsSmallJsonArray(jsonPath);

        // Temporarily adjust the optimization setting for this report.
        bool previousSetting = ReportingEngine.UseReflectionOptimization;
        ReportingEngine.UseReflectionOptimization = !isSmallArray; // Disable if small.

        // Build the report.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "persons");

        // Restore the original global setting.
        ReportingEngine.UseReflectionOptimization = previousSetting;

        // Save the generated report.
        string outputPath = Path.Combine(tempDir, "Report_Output.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated at: {outputPath}");
    }

    // Checks whether the root element of the JSON file is an array whose length
    // is less than or equal to the defined threshold.
    private static bool IsSmallJsonArray(string jsonFilePath)
    {
        using FileStream fs = File.OpenRead(jsonFilePath);
        using JsonDocument doc = JsonDocument.Parse(fs);
        JsonElement root = doc.RootElement;

        // If the root is not an array, treat it as not small.
        if (root.ValueKind != JsonValueKind.Array)
            return false;

        return root.GetArrayLength() <= SmallArrayThreshold;
    }
}
