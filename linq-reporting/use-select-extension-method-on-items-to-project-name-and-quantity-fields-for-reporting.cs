using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSelectExample
{
    // Original data item.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Model that holds the original collection.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Projected item used for reporting.
    public class ProjectedItem
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Wrapper passed as the root object to the reporting engine.
    public class ReportWrapper
    {
        public List<ProjectedItem> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create sample data.
            var model = new ReportModel();
            model.Items.Add(new Item { Name = "Apple", Quantity = 10 });
            model.Items.Add(new Item { Name = "Banana", Quantity = 20 });
            model.Items.Add(new Item { Name = "Cherry", Quantity = 15 });

            // Project the original items to a collection containing only Name and Quantity.
            var projected = model.Items
                .Select(i => new ProjectedItem { Name = i.Name, Quantity = i.Quantity })
                .ToList();

            // Prepare the wrapper that will be used as the root object.
            var wrapper = new ReportWrapper { Items = projected };

            // Create a template document with LINQ Reporting tags.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Name: <<[item.Name]>>, Quantity: <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");
            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Build the report using the wrapper as the data source.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, wrapper, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
