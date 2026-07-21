using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data item.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Value { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            model.Items.Add(new Item { Name = "Alpha",   Value = 10 });
            model.Items.Add(new Item { Name = "Beta",    Value = 20 });
            model.Items.Add(new Item { Name = "Gamma",   Value = 30 });
            model.Items.Add(new Item { Name = "Delta",   Value = 40 });
            model.Items.Add(new Item { Name = "Epsilon", Value = 50 });

            // -----------------------------------------------------------------
            // Create a template document programmatically.
            // The template contains a LINQ Reporting tag that uses ElementAt
            // to fetch the second‑to‑last item from the collection.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Second to last item name: <<[model.Items.ElementAt(model.Items.Count - 2).Name]>>");
            builder.Writeln("Second to last item value: <<[model.Items.ElementAt(model.Items.Count - 2).Value]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back (simulating a real‑world scenario where the template exists on disk).
            var doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            doc.Save(reportPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine("Report generated successfully.");
        }
    }
}
