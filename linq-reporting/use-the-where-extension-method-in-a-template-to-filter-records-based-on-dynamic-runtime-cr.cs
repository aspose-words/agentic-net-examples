using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingWhereExample
{
    // Data model classes
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    public class ReportModel
    {
        // Full collection of products
        public List<Product> Products { get; set; } = new();

        // Runtime criteria – minimum price to include
        public double MinPrice { get; set; }

        // Filtered collection using LINQ Where
        public IEnumerable<Product> FilteredProducts => Products.Where(p => p.Price > MinPrice);
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // LINQ Reporting tags:
            // Loop over the filtered collection defined in the model
            builder.Writeln("<<foreach [p in model.FilteredProducts]>>");
            builder.Writeln("Product: <<[p.Name]>> - Price: $<<[p.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back (required before BuildReport)
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare sample data and runtime filter criteria
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                // Sample products
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 5.0 },
                    new Product { Name = "Banana", Price = 12.0 },
                    new Product { Name = "Cherry", Price = 20.0 },
                    new Product { Name = "Date",   Price = 25.0 }
                },
                // Dynamic filter: include only products priced above 15
                MinPrice = 15.0
            };

            // -------------------------------------------------
            // 4. Build the report using Aspose.Words ReportingEngine
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed

            // The root object name in the template is "model"
            engine.BuildReport(loadedTemplate, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report
            // -------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
