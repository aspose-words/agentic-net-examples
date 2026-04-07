using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Root data model that will be passed to the reporting engine.
    public class ReportModel
    {
        // Collection of items to iterate over in the template.
        public List<Item> Items { get; set; } = new();
    }

    // Simple item class used in the collection.
    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Report of items:");
            // The foreach tag will automatically infer the type of 'item' from the generic List<Item>.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare sample data.
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Index = 1, Name = "Apple" });
            model.Items.Add(new Item { Index = 2, Name = "Banana" });
            model.Items.Add(new Item { Index = 3, Name = "Cherry" });

            // 3. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(template, model, "model");

            // 4. Save the generated report.
            const string outputPath = "ReportOutput.docx";
            template.Save(outputPath);
        }
    }
}
