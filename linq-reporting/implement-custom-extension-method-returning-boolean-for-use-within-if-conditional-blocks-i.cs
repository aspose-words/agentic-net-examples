using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model classes
    public class Item
    {
        public int Value { get; set; } = 0;
    }

    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Extension method used inside the template's if condition
    public static class ItemExtensions
    {
        // Returns true if the item's Value is an even number
        public static bool IsEven(this Item item) => item != null && item.Value % 2 == 0;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Value = 1 },
                    new Item { Value = 2 },
                    new Item { Value = 3 },
                    new Item { Value = 4 }
                }
            };

            // Create a template document programmatically
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // LINQ Reporting tags
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Value]>>");
            builder.Writeln("<<if [item.IsEven()]>> (Even) <</if>>");
            builder.Writeln("<</foreach>>");

            // Save the template
            doc.Save(templatePath);

            // Load the template for reporting
            var template = new Document(templatePath);

            // Configure the reporting engine
            var engine = new ReportingEngine
            {
                // Allow the engine to resolve extension methods and missing members
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Register the type that contains the extension method
            engine.KnownTypes.Add(typeof(ItemExtensions));

            // Build the report using the model as the root object named "model"
            engine.BuildReport(template, model, "model");

            // Save the generated report
            var outputPath = "Report.docx";
            template.Save(outputPath);
        }
    }
}
