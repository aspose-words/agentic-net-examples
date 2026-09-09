using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for the report
    public class Category
    {
        public string Name { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }

    public class ReportModel
    {
        // Optional title used in the template – provide a default value to avoid missing‑member errors.
        public string Title { get; set; } = "Sample Report";

        public List<Category> Categories { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data
            var model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Fruits",
                        Items = new List<Item>
                        {
                            new Item { Name = "Apple",  Price = 0.5m },
                            new Item { Name = "Banana", Price = 0.3m }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Items = new List<Item>
                        {
                            new Item { Name = "Carrot", Price = 0.2m },
                            new Item { Name = "Tomato", Price = 0.4m }
                        }
                    }
                }
            };

            // 2. Create a template document programmatically
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            CreateTemplate(templatePath);

            // 3. Load the template and build the report
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 4. Save the generated report
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
            doc.Save(reportPath);

            Console.WriteLine($"Report generated: {reportPath}");
        }

        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title (optional)
            builder.Writeln("<<[model.Title]>>");

            // Begin outer foreach over categories
            builder.Writeln("<<foreach [category in Categories]>>");
            builder.Writeln("Category: <<[category.Name]>>");
            builder.Writeln(""); // empty line for readability

            // Begin inner foreach over items of the current category
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- <<[item.Name]>> : $<<[item.Price]>>");
            builder.Writeln("<</foreach>>"); // end inner foreach

            builder.Writeln("<</foreach>>"); // end outer foreach

            // Save the template
            doc.Save(filePath);
        }
    }
}
