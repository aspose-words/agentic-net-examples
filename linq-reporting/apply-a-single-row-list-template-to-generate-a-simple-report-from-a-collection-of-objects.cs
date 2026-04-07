using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a collection of products.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Product Report");
            builder.Writeln("<<foreach [p in Products]>>");
            builder.Writeln("- <<[p.Name]>>: $<<[p.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template for reporting.
            var doc = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel();
            model.Products.Add(new Product { Name = "Apple", Price = 1.20 });
            model.Products.Add(new Product { Name = "Banana", Price = 0.80 });
            model.Products.Add(new Product { Name = "Cherry", Price = 2.50 });

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 5. Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
