using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace LinqReportingNestedForeach
{
    // Data model classes
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    public class Category
    {
        public string Name { get; set; } = string.Empty;
        public List<Item> Items { get; set; } = new();

        // Subtotal calculated from Items
        public decimal Subtotal => Items.Sum(i => i.Amount);
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Amount { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // Create template document programmatically
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("LINQ Reporting Example - Subtotals per Category");
            builder.Writeln();

            // Outer foreach over categories
            builder.Writeln("<<foreach [cat in Categories]>>");
            builder.Writeln("Category: <<[cat.Name]>>");
            builder.Writeln();

            // Inner foreach over items of the current category
            builder.Writeln("<<foreach [item in cat.Items]>>");
            // Start a table for each item (simple layout)
            var table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Item");
            builder.InsertCell();
            builder.Writeln("Amount");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Amount]>>");
            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Subtotal for the current category
            builder.Writeln("Subtotal: <<[cat.Subtotal]>>");
            builder.Writeln();
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var reportDoc = new Document(templatePath);

            // Prepare sample data
            var model = new ReportModel
            {
                Categories = new List<Category>
                {
                    new Category
                    {
                        Name = "Beverages",
                        Items = new List<Item>
                        {
                            new Item { Name = "Coffee", Amount = 3.50m },
                            new Item { Name = "Tea", Amount = 2.75m }
                        }
                    },
                    new Category
                    {
                        Name = "Snacks",
                        Items = new List<Item>
                        {
                            new Item { Name = "Chips", Amount = 1.20m },
                            new Item { Name = "Cookies", Amount = 2.10m },
                            new Item { Name = "Nuts", Amount = 3.00m }
                        }
                    }
                }
            };

            // Build the report
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report
            reportDoc.Save(outputPath);
        }
    }
}
