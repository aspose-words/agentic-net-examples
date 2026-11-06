using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report with dynamic background colors:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<backColor [item.Color]>> <<[item.Name]>> <</backColor>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var loadedTemplate = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new() { Name = "Item 1", Color = "LightYellow" },
                new() { Name = "Item 2", Color = "#FFCCCC" },
                new() { Name = "Item 3", Color = "RGB(200,200,255)" }
            }
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        loadedTemplate.Save(outputPath);
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string Color { get; set; } = "";
}
