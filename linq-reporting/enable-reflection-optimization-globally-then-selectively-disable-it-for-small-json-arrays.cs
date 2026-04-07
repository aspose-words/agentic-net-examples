using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting; // JsonDataSource resides in this namespace

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        const string largeJsonPath = "large.json";
        const string smallJsonPath = "small.json";

        // Create a large JSON array (100 items).
        var largeItems = new List<Dictionary<string, object>>();
        for (int i = 1; i <= 100; i++)
        {
            largeItems.Add(new Dictionary<string, object>
            {
                ["Name"] = $"LargeItem{i}",
                ["Value"] = i
            });
        }
        File.WriteAllText(largeJsonPath, System.Text.Json.JsonSerializer.Serialize(largeItems));

        // Create a small JSON array (3 items).
        var smallItems = new List<Dictionary<string, object>>
        {
            new Dictionary<string, object> { ["Name"] = "SmallItemA", ["Value"] = 1 },
            new Dictionary<string, object> { ["Name"] = "SmallItemB", ["Value"] = 2 },
            new Dictionary<string, object> { ["Name"] = "SmallItemC", ["Value"] = 3 }
        };
        File.WriteAllText(smallJsonPath, System.Text.Json.JsonSerializer.Serialize(smallItems));

        // Build template for large data.
        const string largeTemplatePath = "large_template.docx";
        var largeDoc = new Document();
        var largeBuilder = new DocumentBuilder(largeDoc);
        largeBuilder.Writeln("Large Data Report:");
        largeBuilder.Writeln("<<foreach [item in items]>>");
        largeBuilder.Writeln("- <<[item.Name]>> : <<[item.Value]>>");
        largeBuilder.Writeln("<</foreach>>");
        largeDoc.Save(largeTemplatePath);

        // Build template for small data.
        const string smallTemplatePath = "small_template.docx";
        var smallDoc = new Document();
        var smallBuilder = new DocumentBuilder(smallDoc);
        smallBuilder.Writeln("Small Data Report:");
        smallBuilder.Writeln("<<foreach [item in items]>>");
        smallBuilder.Writeln("- <<[item.Name]>> : <<[item.Value]>>");
        smallBuilder.Writeln("<</foreach>>");
        smallDoc.Save(smallTemplatePath);

        // ---------- Report for large JSON array (reflection optimization enabled) ----------
        ReportingEngine.UseReflectionOptimization = true; // Enable globally.

        // Load template and JSON data source.
        var largeTemplate = new Document(largeTemplatePath);
        var largeDataSource = new JsonDataSource(largeJsonPath);

        // Build the report.
        var largeEngine = new ReportingEngine();
        largeEngine.BuildReport(largeTemplate, largeDataSource, "items");

        // Save the generated report.
        const string largeReportPath = "LargeReport.docx";
        largeTemplate.Save(largeReportPath);

        // ---------- Report for small JSON array (reflection optimization disabled) ----------
        ReportingEngine.UseReflectionOptimization = false; // Disable for small collections.

        var smallTemplate = new Document(smallTemplatePath);
        var smallDataSource = new JsonDataSource(smallJsonPath);

        var smallEngine = new ReportingEngine();
        smallEngine.BuildReport(smallTemplate, smallDataSource, "items");

        const string smallReportPath = "SmallReport.docx";
        smallTemplate.Save(smallReportPath);
    }
}
