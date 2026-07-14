using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Sample data model.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Price { get; set; }
    }

    // External type whose static members will be used in the template.
    public static class MyHelper
    {
        public static string FormatPrice(decimal price) => $"${price:F2}";
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // 1. Create the template document programmatically.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // LINQ Reporting tags.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<[item.Name]>> - <<[MyHelper.FormatPrice(item.Price)]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            templateDoc.Save(templatePath);

            // 2. Load the template for reporting.
            Document doc = new Document(templatePath);

            // 3. Prepare a large data set.
            ReportModel model = new ReportModel();
            for (int i = 1; i <= 1000; i++)
            {
                model.Items.Add(new Item
                {
                    Name = $"Product {i}",
                    Price = i * 1.23m
                });
            }

            // 4. Configure the ReportingEngine.
            ReportingEngine.UseReflectionOptimization = true; // Enable reflection optimization.
            ReportingEngine engine = new ReportingEngine();

            // Register the external type so its static members can be used in the template.
            engine.KnownTypes.Add(typeof(MyHelper));

            // 5. Build the report.
            engine.BuildReport(doc, model, "model");

            // 6. Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            doc.Save(reportPath);
        }
    }
}
