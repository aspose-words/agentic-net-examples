using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingContextualAccess
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare folders.
            string workDir = Directory.GetCurrentDirectory();
            string templatePath = Path.Combine(workDir, "template.docx");
            string outputPath = Path.Combine(workDir, "output.docx");

            // -------------------------------------------------
            // 1. Create the template document with LINQ tags.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Product List:");
            // Data band: foreach over Products collection.
            builder.Writeln("<<foreach [p in Products]>>");
            // Use contextual member access (no 'p.' prefix).
            builder.Writeln(" - <<[Name]>> : $<<[Price]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare the data model.
            // -------------------------------------------------
            ReportModel model = new()
            {
                Products = new List<Product>
                {
                    new Product { Name = "Apple", Price = 0.99m },
                    new Product { Name = "Banana", Price = 0.59m },
                    new Product { Name = "Cherry", Price = 2.49m }
                }
            };

            // -------------------------------------------------
            // 4. Build the report.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            doc.Save(outputPath);
        }
    }

    // Root data model.
    public class ReportModel
    {
        public List<Product> Products { get; set; } = new();
    }

    // Item class used inside the data band.
    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
    }
}
