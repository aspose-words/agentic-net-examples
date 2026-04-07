using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Root data model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Order> Orders { get; set; } = new();
    }

    public class Order
    {
        public string CustomerName { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; } // May be null to produce empty paragraphs.
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Orders = new List<Order>
                {
                    new Order
                    {
                        CustomerName = "Alice",
                        Items = new List<Item>
                        {
                            new Item { Name = "Pen", Description = "Blue ballpoint pen" },
                            new Item { Name = "Notebook", Description = null } // No description.
                        }
                    },
                    new Order
                    {
                        CustomerName = "Bob",
                        Items = new List<Item>
                        {
                            new Item { Name = "Mouse", Description = "Wireless mouse" },
                            new Item { Name = "Keyboard", Description = "" } // Empty description.
                        }
                    }
                }
            };

            // 2. Create a template document programmatically.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title.
            builder.Writeln("Orders Report");
            builder.Writeln();

            // Begin foreach over orders.
            builder.Writeln("<<foreach [order in Orders]>>");
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Begin foreach over items of each order.
            builder.Writeln("<<foreach [item in order.Items]>>");
            builder.Writeln("- Item: <<[item.Name]>>");
            // Conditional paragraph that may become empty.
            builder.Writeln("<<if [item.Description != null && item.Description != \"\"]>>Description: <<[item.Description]>> <</if>>");
            builder.Writeln(); // Paragraph separator.
            builder.Writeln("<</foreach>>"); // End items foreach.

            builder.Writeln("<</foreach>>"); // End orders foreach.

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // 3. Load the template for reporting.
            var doc = new Document(templatePath);

            // 4. Configure the reporting engine to remove empty paragraphs.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // 5. Build the report using the model as the root data source.
            engine.BuildReport(doc, model);

            // 6. Save the final document.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
