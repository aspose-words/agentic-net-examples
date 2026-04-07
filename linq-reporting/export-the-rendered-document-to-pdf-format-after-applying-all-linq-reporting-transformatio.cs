using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingPdfExport
{
    // Simple data model for the report.
    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
    }

    // Wrapper class that will be referenced in the template as "model".
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any required encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the final PDF.
            string templatePath = "Template.docx";
            string pdfPath = "Report.pdf";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Add a title.
            builder.Writeln("Product Catalog");
            builder.Writeln();

            // Insert a LINQ Reporting foreach tag to iterate over the Products collection.
            builder.Writeln("<<foreach [product in Products]>>");
            // Each product will be displayed on a separate line.
            builder.Writeln("- <<[product.Name]>> : $<<[product.Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple", Price = 1.20m },
                    new Product { Name = "Banana", Price = 0.80m },
                    new Product { Name = "Cherry", Price = 2.50m }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this simple example.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Export the rendered document to PDF.
            // -----------------------------------------------------------------
            doc.Save(pdfPath, SaveFormat.Pdf);
        }
    }
}
