using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare sample hierarchical data
            var model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Fruits",
                        Items = new List<Item>
                        {
                            new Item { Name = "Apple", Price = 1.20 },
                            new Item { Name = "Banana", Price = 0.80 }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Items = new List<Item>
                        {
                            new Item { Name = "Carrot", Price = 0.50 },
                            new Item { Name = "Tomato", Price = 0.90 }
                        }
                    }
                }
            };

            // Create the template document with LINQ Reporting tags
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Product Catalog");
            builder.Writeln();

            // Outer foreach for categories
            builder.Writeln("<<foreach [category in Categories]>>");
            builder.Writeln("Category: <<[category.Name]>>");
            builder.Writeln();

            // Inner foreach for items within each category
            builder.Writeln("<<foreach [item in category.Items]>>");
            builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a file
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template for report generation
            var doc = new Document(templatePath);

            // Build the report using the data model
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final report
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }

    // Root data model
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    // Category containing a collection of items
    public class Category
    {
        public string Name { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    // Individual item
    public class Item
    {
        public string Name { get; set; } = "";
        public double Price { get; set; }
    }
}
