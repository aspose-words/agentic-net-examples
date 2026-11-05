using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingNestedLoops
{
    // Data model classes
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Category
    {
        public string Name { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
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
                            new Item { Name = "Apple", Quantity = 10 },
                            new Item { Name = "Banana", Quantity = 20 }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Items = new List<Item>
                        {
                            new Item { Name = "Carrot", Quantity = 15 },
                            new Item { Name = "Tomato", Quantity = 12 }
                        }
                    }
                }
            };

            // Create a new blank document and a builder for the template
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Build the LINQ Reporting template with nested foreach loops
            builder.Writeln("<<foreach [category in Categories]>>");
            builder.Writeln("Category: <<[category.Name]>>");
            builder.Writeln("<<foreach [item in category.Items]>>");
            builder.Writeln("- <<[item.Name]>> (Qty: <<[item.Quantity]>>)");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Generate the report using the ReportingEngine
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated document
            doc.Save("Report.docx");
        }
    }
}
