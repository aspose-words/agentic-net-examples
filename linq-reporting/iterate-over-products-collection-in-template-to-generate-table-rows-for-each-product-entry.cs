using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple product model.
    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
    }

    // Wrapper model that contains the collection referenced by the template.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document with LINQ Reporting tags.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Table header.
            builder.Writeln("Product Report");
            builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();
            builder.EndTable();

            // Begin foreach loop over the Products collection.
            builder.Writeln("<<foreach [p in Products]>>");

            // Table rows generated for each product.
            var table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Price]>>");
            builder.EndRow();
            builder.EndTable();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template (required before building the report).
            var reportDoc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 0.99m },
                    new Product { Name = "Banana", Price = 0.59m },
                    new Product { Name = "Cherry", Price = 2.49m }
                }
            };

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
