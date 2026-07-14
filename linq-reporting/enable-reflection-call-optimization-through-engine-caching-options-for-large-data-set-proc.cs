using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Model classes used by the LINQ Reporting engine.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Order
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a large data set (e.g., 10,000 items).
            var order = new Order();
            for (int i = 1; i <= 10000; i++)
            {
                order.Items.Add(new Item
                {
                    Index = i,
                    Name = $"Product #{i}"
                });
            }

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Order Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            doc.Save(templatePath);

            // Load the template (demonstrates load step).
            var loadedDoc = new Document(templatePath);

            // Enable reflection optimization (engine caching) for large data sets.
            ReportingEngine.UseReflectionOptimization = true;

            // Create the reporting engine and optionally set options.
            var engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };

            // Build the report using the root object name "order".
            engine.BuildReport(loadedDoc, order, "order");

            // Save the generated report.
            var outputPath = "Report.docx";
            loadedDoc.Save(outputPath);
        }
    }
}
