using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model for a product
    public class Product
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
        public bool InStock { get; set; }
    }

    // Wrapper class that will be passed as the root data source
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -------------------------
            // 1. Create the template document
            // -------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a foreach loop over the Products collection
            builder.Writeln("<<foreach [p in Products]>>");
            // Display product name
            builder.Writeln("Product: <<[p.Name]>>");
            // If the product is in stock, show the quantity; otherwise show "Out of stock"
            builder.Writeln("<<if [p.InStock]>>Quantity: <<[p.Quantity]>> <<else>>Out of stock<</if>>");
            // End the foreach loop
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before loading for reporting)
            const string templatePath = "ReportTemplate.docx";
            template.Save(templatePath);

            // -------------------------
            // 2. Load the template for report generation
            // -------------------------
            Document doc = new Document(templatePath);

            // -------------------------
            // 3. Prepare sample data
            // -------------------------
            ReportModel model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple",  Quantity = 10, InStock = true },
                    new Product { Name = "Banana", Quantity = 0,  InStock = false },
                    new Product { Name = "Cherry", Quantity = 5,  InStock = true }
                }
            };

            // -------------------------
            // 4. Build the report using the LINQ Reporting engine
            // -------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------
            // 5. Save the generated report
            // -------------------------
            const string resultPath = "ReportResult.docx";
            doc.Save(resultPath);
        }
    }
}
