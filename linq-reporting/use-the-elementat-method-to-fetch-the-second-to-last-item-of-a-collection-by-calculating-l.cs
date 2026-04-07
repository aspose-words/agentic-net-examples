using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Collection of items; initialized to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item with a name and a numeric value.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that uses ElementAt to fetch the
            // second‑to‑last item from the collection (Count - 2) and displays its Name.
            builder.Writeln("Second to last item: <<[model.Items.ElementAt(model.Items.Count - 2).Name]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data for the report.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();
            model.Items.Add(new Item { Name = "Alpha",   Value = 10 });
            model.Items.Add(new Item { Name = "Bravo",   Value = 20 });
            model.Items.Add(new Item { Name = "Charlie", Value = 30 });
            model.Items.Add(new Item { Name = "Delta",   Value = 40 });
            model.Items.Add(new Item { Name = "Echo",    Value = 50 });

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using the data model.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // BuildReport overload that specifies the root name ("model") used in the template.
            engine.BuildReport(loadedTemplate, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string reportPath = "Report.docx";
            loadedTemplate.Save(reportPath);
        }
    }
}
