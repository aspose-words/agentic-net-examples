using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Entry
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class ReportModel
{
    public List<Entry> FilteredEntries { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample JSON data
        string jsonPath = "data.json";
        var sampleData = new List<Entry>
        {
            new Entry { Id = 1, Name = "Laptop", Category = "Electronics" },
            new Entry { Id = 2, Name = "Desk", Category = "Furniture" },
            new Entry { Id = 3, Name = "Smartphone", Category = "electronics" },
            new Entry { Id = 4, Name = "Chair", Category = "Furniture" },
            new Entry { Id = 5, Name = "Headphones", Category = "ELECTRONICS" }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON and filter case‑insensitively on Category
        var allEntries = JsonConvert.DeserializeObject<List<Entry>>(File.ReadAllText(jsonPath)) ?? new List<Entry>();
        var filtered = allEntries
            .Where(e => e.Category.Equals("electronics", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Prepare the root model for the report
        var model = new ReportModel { FilteredEntries = filtered };

        // Create the template document
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Filtered Entries (Category = Electronics):");
        builder.Writeln("<<foreach [e in FilteredEntries]>>");
        builder.Writeln("- ID: <<[e.Id]>>, Name: <<[e.Name]>>, Category: <<[e.Category]>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template as required by the workflow
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);
        var loadedTemplate = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the final report
        string outputPath = "report.docx";
        loadedTemplate.Save(outputPath);
    }
}
