using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Root data model that will be passed to the reporting engine.
    public class ReportModel
    {
        // The generic list whose item type will be inferred automatically by the engine.
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class used in the generic list.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Items.Add(new Item { Index = 1, Name = "Apple" });
            model.Items.Add(new Item { Index = 2, Name = "Banana" });
            model.Items.Add(new Item { Index = 3, Name = "Cherry" });

            // Create a new blank document that will serve as the template.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Write static text.
            builder.Writeln("Item List:");
            // Insert a foreach tag that iterates over the generic list.
            // The engine will infer that each element is of type 'Item'.
            builder.Writeln("<<foreach [item in Items]>>");
            // Inside the loop we can reference the properties of each inferred item.
            builder.Writeln("  Index: <<[item.Index]>>, Name: <<[item.Name]>>");
            // Close the foreach block.
            builder.Writeln("<</foreach>>");

            // Build the report using the model as the root data source.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            // The root object name must match the name used in the template ("model").
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
