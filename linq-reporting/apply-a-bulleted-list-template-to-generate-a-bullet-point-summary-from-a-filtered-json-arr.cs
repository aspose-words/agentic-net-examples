using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public string Name { get; set; } = "";
    public bool IsImportant { get; set; }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data.
        string json = @"[
            { ""Name"": ""Item A"", ""IsImportant"": true },
            { ""Name"": ""Item B"", ""IsImportant"": false },
            { ""Name"": ""Item C"", ""IsImportant"": true }
        ]";

        // Deserialize JSON into a list of items.
        List<Item> allItems = JsonConvert.DeserializeObject<List<Item>>(json) ?? new();

        // Filter items that are marked as important.
        List<Item> filteredItems = allItems.Where(i => i.IsImportant).ToList();

        // Prepare the model for the report.
        ReportModel model = new()
        {
            Items = filteredItems
        };

        // Create the LINQ Reporting template programmatically.
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Important Items Summary:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("• <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
