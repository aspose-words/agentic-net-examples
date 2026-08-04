using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Lists;

namespace AsposeWordsLinqReportingExample
{
    // Data entity representing a single item.
    public class Item
    {
        public string Category { get; set; } = "";
        public string Name { get; set; } = "";
    }

    // Wrapper for a group of items.
    public class Group
    {
        public string Category { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<Group> Groups { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for .NET Core.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Prepare sample data.
            // -----------------------------------------------------------------
            var items = new List<Item>
            {
                new Item { Category = "Fruits", Name = "Apple" },
                new Item { Category = "Fruits", Name = "Banana" },
                new Item { Category = "Fruits", Name = "Cherry" },
                new Item { Category = "Vegetables", Name = "Carrot" },
                new Item { Category = "Vegetables", Name = "Lettuce" },
                new Item { Category = "Grains", Name = "Rice" }
            };

            // Group items by Category and map to Group objects.
            var model = new ReportModel
            {
                Groups = items
                    .GroupBy(i => i.Category)
                    .Select(g => new Group
                    {
                        Category = g.Key,
                        Items = g.ToList()
                    })
                    .ToList()
            };

            // -----------------------------------------------------------------
            // 2. Create a template document programmatically.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Create a bullet list template that will be used for both levels.
            List bulletList = doc.Lists.Add(ListTemplate.BulletDefault);

            // Outer foreach over groups.
            builder.Writeln("<<foreach [group in Groups]>>");

            // Apply first‑level bullet for the group name.
            builder.ListFormat.List = bulletList;
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("<<[group.Category]>>");

            // Inner foreach over items within the current group.
            builder.Writeln("<<foreach [item in group.Items]>>");

            // Apply second‑level bullet for each item name.
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("<<[item.Name]>>");

            // Close inner foreach.
            builder.Writeln("<</foreach>>");

            // Close outer foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var template = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model; the root name is "model".
            engine.BuildReport(template, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            template.Save(outputPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
