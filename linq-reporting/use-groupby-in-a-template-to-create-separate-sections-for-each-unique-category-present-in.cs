using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Product
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
    public double Price { get; set; }
}

public class CategoryGroup
{
    public string Category { get; set; } = "";
    public List<Product> Items { get; set; } = new();
}

public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // 1. Prepare sample data.
        var products = new List<Product>
        {
            new() { Name = "Apple",    Category = "Fruit",      Price = 1.20 },
            new() { Name = "Banana",   Category = "Fruit",      Price = 0.80 },
            new() { Name = "Carrot",   Category = "Vegetable", Price = 0.50 },
            new() { Name = "Broccoli", Category = "Vegetable", Price = 0.90 }
        };

        // 2. Group products by Category.
        var model = new ReportModel
        {
            Groups = products
                .GroupBy(p => p.Category)
                .Select(g => new CategoryGroup
                {
                    Category = g.Key,
                    Items = g.ToList()
                })
                .ToList()
        };

        // 3. Create the LINQ Reporting template.
        string templatePath = Path.Combine(workDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer foreach over groups.
        builder.Writeln("<<foreach [group in data.Groups]>>");
        builder.Writeln("Category: <<[group.Category]>>");
        // Inner foreach over items in each group.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 4. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 5. Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "data");

        // 6. Save the generated report.
        string outputPath = Path.Combine(workDir, "Report.docx");
        reportDoc.Save(outputPath);
    }
}
