using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a collection of items.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Items.Add(new Item { Name = "Apple" });
            model.Items.Add(new Item { Name = "Banana" });
            model.Items.Add(new Item { Name = "Cherry" });

            // -----------------------------------------------------------------
            // Step 1: Create a template document with a LINQ Reporting tag.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // The tag uses ElementAt together with Count to fetch the last item.
            builder.Writeln("Last item in collection: <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // Step 2: Load the template and build the report.
            // -----------------------------------------------------------------
            var template = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model; the root name is "model".
            engine.BuildReport(template, model, "model");

            // -----------------------------------------------------------------
            // Step 3: Save the generated report.
            // -----------------------------------------------------------------
            var outputPath = "Report.docx";
            template.Save(outputPath);

            // The program finishes without waiting for user input.
        }
    }
}
