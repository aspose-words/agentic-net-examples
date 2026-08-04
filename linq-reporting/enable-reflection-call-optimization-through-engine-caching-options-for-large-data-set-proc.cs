using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Data model for the report.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var templatePath = "Template.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Simple foreach loop that will list all items.
        builder.Writeln("Report of Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        var templateDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare a large data set.
        // -----------------------------------------------------------------
        var model = new ReportModel();

        const int itemCount = 10000; // Simulate a large collection.
        for (int i = 1; i <= itemCount; i++)
        {
            model.Items.Add(new Item
            {
                Id = i,
                Name = $"Item #{i}"
            });
        }

        // -----------------------------------------------------------------
        // 4. Enable reflection optimization (engine caching) and build the report.
        // -----------------------------------------------------------------
        ReportingEngine.UseReflectionOptimization = true; // Enable caching of reflection calls.

        var engine = new ReportingEngine();
        // No special options are required for this scenario, but the property is set explicitly.
        engine.Options = ReportBuildOptions.None;

        // Build the report using the root object name "model" to match the tags.
        engine.BuildReport(templateDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        var outputPath = "ReportOutput.docx";
        templateDoc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {outputPath}");
    }
}
