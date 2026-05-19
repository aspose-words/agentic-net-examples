using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Needed for the Table class

namespace LinqReportingExample
{
    // Data model for a product.
    public class Product
    {
        public string Name { get; set; } = "";
        public double Price { get; set; }
    }

    // Root model containing the collection used in the template.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Product Report");
            builder.Writeln(""); // Empty line for spacing.

            // Begin foreach loop over the Products collection.
            builder.Writeln("<<foreach [p in Products]>>");

            // Create a table for each product.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Data row – values will be filled by the reporting engine.
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Price]>>");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // End foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 1.20 },
                    new Product { Name = "Banana", Price = 0.80 },
                    new Product { Name = "Orange", Price = 1.50 }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // BuildReport uses the model as the root data source; members can be accessed directly.
            engine.BuildReport(reportDoc, model);

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
