using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingConditionalExample
{
    // Data model classes
    public class ReportModel
    {
        public decimal Threshold { get; set; } = 0m;
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        public decimal Amount { get; set; } = 0m;
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // Create the template document with LINQ Reporting tags
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report of Items");
            builder.Writeln("Threshold: <<[model.Threshold]>>");
            builder.Writeln();

            // Begin foreach over Items
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Item: <<[item.Name]>>");
            builder.Writeln("Amount: <<[item.Amount]>>");

            // Conditional section – appears only when the amount exceeds the model threshold
            builder.Writeln("<<if [item.Amount > model.Threshold]>>");
            builder.Writeln("**High value item!**");
            builder.Writeln("<</if>>");

            // End foreach
            builder.Writeln("<</foreach>>");

            // Save the template to disk
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template for report generation
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // Sample data
            ReportModel model = new()
            {
                Threshold = 1000m,
                Items = new()
                {
                    new Item { Name = "Item A", Amount = 500m },
                    new Item { Name = "Item B", Amount = 1500m },
                    new Item { Name = "Item C", Amount = 2000m }
                }
            };

            // Build the report using LINQ Reporting Engine
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report
            doc.Save(reportPath);
        }
    }
}
