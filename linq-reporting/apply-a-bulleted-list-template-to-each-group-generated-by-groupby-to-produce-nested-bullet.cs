using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingBulletedGroups
{
    // Data model classes
    public class Item
    {
        public string Category { get; set; } = "";
        public string Name { get; set; } = "";
    }

    public class CategoryGroup
    {
        public string Category { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class ReportModel
    {
        public List<CategoryGroup> Groups { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var items = new List<Item>
            {
                new() { Category = "Fruits", Name = "Apple" },
                new() { Category = "Fruits", Name = "Banana" },
                new() { Category = "Fruits", Name = "Orange" },
                new() { Category = "Vegetables", Name = "Carrot" },
                new() { Category = "Vegetables", Name = "Broccoli" },
                new() { Category = "Grains", Name = "Rice" },
                new() { Category = "Grains", Name = "Wheat" }
            };

            // Group items by Category
            var model = new ReportModel
            {
                Groups = items
                    .GroupBy(i => i.Category)
                    .Select(g => new CategoryGroup
                    {
                        Category = g.Key,
                        Items = g.ToList()
                    })
                    .ToList()
            };

            // Create the template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Outer foreach over groups
            builder.Writeln("<<foreach [group in model.Groups]>>");

            // Bullet for group name
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("<<[group.Category]>>");

            // Indent for nested items
            builder.ListFormat.ListIndent();

            // Inner foreach over items in the current group
            builder.Writeln("<<foreach [item in group.Items]>>");
            builder.Writeln("<<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Outdent back to outer level
            builder.ListFormat.ListOutdent();

            // End outer foreach
            builder.Writeln("<</foreach>>");

            // Clean up list formatting
            builder.ListFormat.RemoveNumbers();

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
