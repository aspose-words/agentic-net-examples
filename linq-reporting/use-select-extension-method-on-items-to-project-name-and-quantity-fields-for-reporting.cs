using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSelectExample
{
    // Sample data class with extra fields.
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    // Projection class containing only the fields needed for the report.
    public class ItemProjection
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<ItemProjection> ProjectedItems { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            List<Item> items = new()
            {
                new Item { Name = "Apple", Quantity = 10, Price = 0.5m },
                new Item { Name = "Banana", Quantity = 20, Price = 0.3m },
                new Item { Name = "Cherry", Quantity = 15, Price = 1.2m }
            };

            // Project Name and Quantity using LINQ Select.
            ReportModel model = new()
            {
                ProjectedItems = items
                    .Select(i => new ItemProjection { Name = i.Name, Quantity = i.Quantity })
                    .ToList()
            };

            // Create the template document.
            string templatePath = "Template.docx";
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Items Report");
            builder.Writeln("==============");
            builder.Writeln("<<foreach [p in ProjectedItems]>>");
            builder.Writeln("- <<[p.Name]>>: <<[p.Quantity]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template and build the report.
            Document reportDoc = new(templatePath);
            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
