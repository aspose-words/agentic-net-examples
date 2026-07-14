using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model for the report.
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    // Represents a category that contains a collection of items.
    public class Category
    {
        public string Name { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    // Represents an individual item.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Categories Report:");
            builder.Writeln("<<foreach [category in Categories]>>");
            builder.Writeln("Category: <<[category.Name]>>");
            builder.Writeln("Items:");
            builder.Writeln("<<foreach [item in category.Items]>>");
            builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            var document = new Document(templatePath);

            var model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Fruits",
                        Items = new List<Item>
                        {
                            new Item { Index = 1, Name = "Apple" },
                            new Item { Index = 2, Name = "Banana" },
                            new Item { Index = 3, Name = "Cherry" }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Items = new List<Item>
                        {
                            new Item { Index = 1, Name = "Carrot" },
                            new Item { Index = 2, Name = "Lettuce" }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(document, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            document.Save(outputPath);
        }
    }
}
