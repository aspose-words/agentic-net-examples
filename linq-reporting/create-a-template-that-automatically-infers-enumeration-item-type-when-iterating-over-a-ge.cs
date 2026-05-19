using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model classes
    public class ReportModel
    {
        // The engine will infer the item type (Item) when iterating over this list.
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Report generated with Aspose.Words LINQ Reporting:");
            // The foreach tag iterates over the generic list 'Items'.
            // The engine automatically infers that each element is of type 'Item'.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("- Name: <<[item.Name]>>");
            builder.Writeln("- Quantity: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Name = "Apple", Quantity = 10 });
            model.Items.Add(new Item { Name = "Banana", Quantity = 7 });
            model.Items.Add(new Item { Name = "Cherry", Quantity = 15 });

            // -----------------------------------------------------------------
            // 4. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportOutput.docx";
            reportDoc.Save(outputPath);
        }
    }
}
