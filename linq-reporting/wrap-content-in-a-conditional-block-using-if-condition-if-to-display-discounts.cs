using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class ReportModel
    {
        // Collection of items to be displayed in the report
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
        public decimal Discount { get; set; }

        // Helper property used in the conditional tag
        public bool HasDiscount => Discount > 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create the template document programmatically
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Header
            builder.Writeln("Item Report");
            builder.Writeln();

            // Begin foreach loop over Items
            builder.Writeln("<<foreach [item in Items]>>");

            // Item name and price
            builder.Writeln("Name: <<[item.Name]>>");
            builder.Writeln("Price: $<<[item.Price]>>");

            // Conditional block: display discount only when it exists
            builder.Writeln("<<if [item.HasDiscount]>>Discount: $<<[item.Discount]>> <</if>>");

            // Add a separator line between items
            builder.Writeln("--------------------");

            // End foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to a file (required before building the report)
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template document
            Document reportDoc = new Document(templatePath);

            // Step 3: Prepare sample data
            ReportModel model = new()
            {
                Items = new List<Item>
                {
                    new Item { Name = "Laptop", Price = 1200m, Discount = 150m },
                    new Item { Name = "Smartphone", Price = 800m, Discount = 0m },
                    new Item { Name = "Headphones", Price = 150m, Discount = 20m }
                }
            };

            // Step 4: Build the report using the LINQ Reporting engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Step 5: Save the generated report
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
