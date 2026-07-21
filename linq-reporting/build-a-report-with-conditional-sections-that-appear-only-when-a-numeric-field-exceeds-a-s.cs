using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = "Untitled Report";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public double Value { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Report title.
        builder.Writeln("Report: <<[model.Title]>>");
        builder.Writeln();

        // Begin a foreach loop over the collection of items.
        builder.Writeln("<<foreach [item in model.Items]>>");

        // Item details.
        builder.Writeln("Item: <<[item.Name]>> - Value: <<[item.Value]>>");

        // Conditional section: appears only when the numeric value exceeds 50.
        builder.Writeln("<<if [item.Value > 50]>>");
        builder.Writeln("  *** High value detected! ***");
        builder.Writeln("<</if>>");

        // End of the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Sample data model.
        ReportModel model = new()
        {
            Title = "Quarterly Sales Report",
            Items = new List<Item>
            {
                new() { Name = "Product A", Value = 42.5 },
                new() { Name = "Product B", Value = 67.0 },
                new() { Name = "Product C", Value = 15.3 },
                new() { Name = "Product D", Value = 89.9 }
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}
