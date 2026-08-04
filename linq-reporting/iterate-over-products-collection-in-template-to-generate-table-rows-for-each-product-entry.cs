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
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }

    // Wrapper model that contains the collection used in the template.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin the foreach loop over the Products collection.
            builder.Writeln("<<foreach [p in model.Products]>>");

            // Create a table header.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Product Name");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Row that will be repeated for each product.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Price]>>");
            builder.EndRow();

            // Finish the table and the foreach block.
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple", Price = 1.23m },
                    new Product { Name = "Banana", Price = 0.99m },
                    new Product { Name = "Cherry", Price = 2.50m }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(outputPath);
        }
    }
}
