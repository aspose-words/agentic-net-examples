using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for encodings required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplate(templatePath);

        // -----------------------------------------------------------------
        // 2. Create sample JSON data files.
        // -----------------------------------------------------------------
        string largeJsonPath = Path.Combine(outputDir, "data_large.json");
        string smallJsonPath = Path.Combine(outputDir, "data_small.json");
        CreateJsonData(largeJsonPath, 100); // large array (100 items)
        CreateJsonData(smallJsonPath, 2);   // small array (2 items)

        // -----------------------------------------------------------------
        // 3. Generate report using large JSON (optimization stays enabled).
        // -----------------------------------------------------------------
        string largeReportPath = Path.Combine(outputDir, "ReportLarge.docx");
        GenerateReport(templatePath, largeJsonPath, largeReportPath);

        // -----------------------------------------------------------------
        // 4. Generate report using small JSON (disable optimization for this run).
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = false; // temporarily disable
        string smallReportPath = Path.Combine(outputDir, "ReportSmall.docx");
        GenerateReport(templatePath, smallJsonPath, smallReportPath);

        // Reset to default (optional).
        ReportingEngine.UseReflectionOptimization = true;
    }

    // Creates a simple Word template with a foreach loop over the JSON root collection.
    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Sample Report");
        // The JSON root is treated as a collection, so iterate directly over it.
        builder.Writeln("<<foreach [item in data]>>");
        builder.Writeln("Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    // Generates a JSON file containing an object with an "Items" array.
    private static void CreateJsonData(string filePath, int itemCount)
    {
        var items = new List<Dictionary<string, object>>();
        for (int i = 1; i <= itemCount; i++)
        {
            items.Add(new Dictionary<string, object>
            {
                { "Name", $"Item{i}" },
                { "Value", i * 10 }
            });
        }

        var root = new Dictionary<string, object>
        {
            { "Items", items }
        };

        string json = System.Text.Json.JsonSerializer.Serialize(
            root,
            new System.Text.Json.JsonSerializerOptions { WriteIndented = true });

        File.WriteAllText(filePath, json, Encoding.UTF8);
    }

    // Loads the template, binds the JSON data source, and builds the report.
    private static void GenerateReport(string templatePath, string jsonPath, string outputPath)
    {
        // Load the template document.
        Document template = new Document(templatePath);

        // Create a JsonDataSource from the JSON file.
        JsonDataSource dataSource = new JsonDataSource(jsonPath);

        // Build the report. The root object name used in the template tags is "data".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "data");

        // Save the generated report.
        template.Save(outputPath);
    }
}
