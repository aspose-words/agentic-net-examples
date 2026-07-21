using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSelectExample
{
    // Simple data entity.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
        // Additional fields can be added without affecting the projection.
    }

    // Projection used for the report.
    public class ItemProjection
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Wrapper object passed to the reporting engine.
    public class ReportData
    {
        public List<ItemProjection> ProjectedItems { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [item in ProjectedItems]>>");
            builder.Writeln("- <<[item.Name]>> : <<[item.Quantity]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template back (required before building the report).
            var doc = new Document(templatePath);

            // 3. Prepare source data.
            var items = new List<Item>
            {
                new Item { Name = "Apple", Quantity = 10 },
                new Item { Name = "Banana", Quantity = 20 },
                new Item { Name = "Orange", Quantity = 15 }
            };

            // 4. Project the required fields using LINQ Select.
            var projected = items
                .Select(i => new ItemProjection { Name = i.Name, Quantity = i.Quantity })
                .ToList();

            // 5. Wrap the projection for the reporting engine.
            var reportData = new ReportData { ProjectedItems = projected };

            // 6. Build the report.
            var engine = new ReportingEngine();
            // No special options are needed for this simple scenario.
            engine.BuildReport(doc, reportData, "");

            // 7. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
