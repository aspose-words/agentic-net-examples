using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare the template document.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Create sample JSON data files.
        const string largeJsonPath = "large.json";
        const string smallJsonPath = "small.json";
        CreateJsonFile(largeJsonPath, 100); // Large array.
        CreateJsonFile(smallJsonPath, 2);   // Small array.

        // Generate report for large JSON (optimization stays enabled).
        GenerateReport(templatePath, largeJsonPath, "items", "LargeReport.docx");

        // Disable reflection optimization for small JSON arrays.
        ReportingEngine.UseReflectionOptimization = false;

        // Generate report for small JSON (optimization disabled).
        GenerateReport(templatePath, smallJsonPath, "items", "SmallReport.docx");

        // Reset to default if needed.
        ReportingEngine.UseReflectionOptimization = true;
    }

    // Creates a simple template with a foreach loop over a collection named "items".
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use static DateTime.Now; add DateTime to known types at runtime.
        builder.Writeln("Report generated on: <<[DateTime.Now]>>");
        builder.Writeln();
        builder.Writeln("<<foreach [item in items]>>");
        builder.Writeln("Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Generates a JSON file containing an array of objects with Name and Value properties.
    private static void CreateJsonFile(string path, int count)
    {
        var items = new List<Dictionary<string, object>>();
        for (int i = 1; i <= count; i++)
        {
            items.Add(new Dictionary<string, object>
            {
                ["Name"] = $"Item{i}",
                ["Value"] = i * 10
            });
        }

        string json = System.Text.Json.JsonSerializer.Serialize(items);
        File.WriteAllText(path, json);
    }

    // Builds a report using the specified template and JSON data source.
    private static void GenerateReport(string templatePath, string jsonPath, string rootName, string outputPath)
    {
        // Load the template.
        Document doc = new Document(templatePath);

        // Create a JSON data source.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();

        // Register DateTime type to allow static member access in the template.
        engine.KnownTypes.Add(typeof(DateTime));

        // Use the overload that specifies the root name.
        engine.BuildReport(doc, dataSource, rootName);

        // Save the generated report.
        doc.Save(outputPath);
    }
}
