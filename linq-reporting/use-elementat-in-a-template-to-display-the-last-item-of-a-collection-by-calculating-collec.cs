using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model with a collection of items.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();

        // Populate sample data in the constructor.
        public ReportModel()
        {
            Items.Add(new Item { Name = "Alpha", Value = 10 });
            Items.Add(new Item { Name = "Beta", Value = 20 });
            Items.Add(new Item { Name = "Gamma", Value = 30 });
        }
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for .NET Core environments.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Step 1: Create the template document with a LINQ Reporting tag that
            // uses ElementAt to fetch the last item based on collection length.
            var template = new Document();
            var builder = new DocumentBuilder(template);
            builder.Writeln("Last item name: <<[model.Items.ElementAt(model.Items.Count - 1).Name]>>");
            builder.Writeln("Last item value: <<[model.Items.ElementAt(model.Items.Count - 1).Value]>>");
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Step 2: Load the template for report generation.
            var doc = new Document(templatePath);

            // Step 3: Prepare the data source.
            var model = new ReportModel();

            // Step 4: Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Step 5: Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
