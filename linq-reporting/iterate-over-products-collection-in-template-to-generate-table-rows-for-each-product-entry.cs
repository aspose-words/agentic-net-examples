using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingExample
{
    // Data model for a product.
    public class Product
    {
        public int Index { get; set; } = 0;
        public string Name { get; set; } = "";
        public decimal Price { get; set; } = 0m;
    }

    // Root model containing the collection of products.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // ---------- Create the template ----------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Product Report");
            builder.Writeln("<<foreach [p in Products]>>");

            // Start a table inside the foreach block.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Data row – will be repeated for each product.
            builder.InsertCell();
            builder.Writeln("<<[p.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Price]>>");
            builder.EndRow();

            // End the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // ---------- Prepare sample data ----------
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Index = 1, Name = "Apple",  Price = 0.50m },
                    new Product { Index = 2, Name = "Banana", Price = 0.30m },
                    new Product { Index = 3, Name = "Cherry", Price = 0.20m }
                }
            };

            // ---------- Build the report ----------
            Document report = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            engine.BuildReport(report, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            report.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated: {reportPath}");
        }
    }
}
