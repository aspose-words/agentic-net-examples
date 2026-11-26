using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingSelectExample
{
    // Source data model
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
        public decimal Price { get; set; }
    }

    // Projection used in the report
    public class ItemProjection
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }

    // Root object passed to the reporting engine
    public class ReportModel
    {
        public List<ItemProjection> ProjectedItems { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Create sample source items
            List<Item> items = new()
            {
                new Item { Name = "Apple", Quantity = 10, Price = 0.5m },
                new Item { Name = "Banana", Quantity = 20, Price = 0.3m },
                new Item { Name = "Orange", Quantity = 15, Price = 0.4m }
            };

            // Project only Name and Quantity using LINQ Select
            ReportModel model = new()
            {
                ProjectedItems = items
                    .Select(i => new ItemProjection { Name = i.Name, Quantity = i.Quantity })
                    .ToList()
            };

            // Build the LINQ Reporting template programmatically
            const string templatePath = "Template.docx";
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Items Report");
            builder.Writeln("<<foreach [p in ProjectedItems]>>");
            builder.Writeln("Name: <<[p.Name]>>, Quantity: <<[p.Quantity]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template and generate the report
            Document reportDoc = new(templatePath);
            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
