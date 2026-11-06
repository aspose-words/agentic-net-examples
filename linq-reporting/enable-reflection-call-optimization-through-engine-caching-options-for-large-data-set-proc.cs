using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingReflectionOptimization
{
    // Simple product model.
    public class Product
    {
        public int Index { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    // Wrapper class that holds the collection used by the report and additional metadata.
    public class ReportData
    {
        public List<Product> Products { get; set; } = new();
        public string GeneratedOn { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any encoding needs.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare output folder.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Paths for the template and the generated report.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            string resultPath = Path.Combine(outputDir, "Report.docx");

            // -----------------------------------------------------------------
            // Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Product Report");
            builder.Writeln("Generated on: <<[GeneratedOn]>>");
            builder.Writeln();

            // Begin foreach over Products.
            builder.Writeln("<<foreach [p in Products]>>");

            // Create a table header.
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Row that will be repeated for each product.
            builder.InsertCell();
            builder.Writeln("<<[p.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[p.Name]>>");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // End foreach.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and generate the report.
            // -----------------------------------------------------------------
            Document doc = new(templatePath);

            // Enable reflection optimization (engine caching).
            ReportingEngine.UseReflectionOptimization = true;

            // Create a large data set.
            ReportData data = new();
            data.GeneratedOn = DateTime.Now.ToString("g");
            for (int i = 1; i <= 1000; i++)
            {
                data.Products.Add(new Product
                {
                    Index = i,
                    Name = $"Product #{i}"
                });
            }

            // Build the report.
            ReportingEngine engine = new();
            engine.BuildReport(doc, data, "data");

            // Save the generated report.
            doc.Save(resultPath);

            Console.WriteLine($"Report generated: {resultPath}");
        }
    }
}
