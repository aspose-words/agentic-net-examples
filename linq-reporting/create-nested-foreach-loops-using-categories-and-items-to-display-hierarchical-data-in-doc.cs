using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of categories.
        public List<Category> Categories { get; set; } = new();
    }

    public class Category
    {
        // Category name.
        public string Name { get; set; } = string.Empty;

        // Items belonging to this category.
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        // Item name.
        public string Name { get; set; } = string.Empty;

        // Quantity of the item.
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Categories =
                {
                    new Category
                    {
                        Name = "Fruits",
                        Items =
                        {
                            new Item { Name = "Apple", Quantity = 10 },
                            new Item { Name = "Banana", Quantity = 5 }
                        }
                    },
                    new Category
                    {
                        Name = "Vegetables",
                        Items =
                        {
                            new Item { Name = "Carrot", Quantity = 7 },
                            new Item { Name = "Tomato", Quantity = 12 }
                        }
                    }
                }
            };

            // -----------------------------------------------------------------
            // 2. Create the template document programmatically.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Hierarchical Report");
            builder.Writeln();

            // Outer foreach over categories.
            builder.Writeln("<<foreach [category in Categories]>>");
            builder.Writeln("Category: <<[category.Name]>>");
            builder.Writeln("Items:");

            // Inner foreach over items of the current category.
            builder.Writeln("<<foreach [item in category.Items]>>");
            builder.Writeln("- <<[item.Name]>> (Qty: <<[item.Quantity]>>)");
            builder.Writeln("<</foreach>>"); // End inner foreach.

            builder.Writeln("<</foreach>>"); // End outer foreach.

            // Save the template to a temporary file.
            const string templatePath = "ReportTemplate.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root data source named "model".
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);

            // The program finishes without waiting for user input.
        }
    }
}
