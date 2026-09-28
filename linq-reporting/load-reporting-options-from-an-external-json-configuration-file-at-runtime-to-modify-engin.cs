using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace LinqReportingOptionsExample
{
    // Model classes
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public List<Item> Items { get; set; } = new()
        {
            new Item { Index = 1, Name = "Apple", Price = 0.5 },
            new Item { Index = 2, Name = "Banana", Price = 0.3 },
            new Item { Index = 3, Name = "Cherry", Price = 0.8 }
        };
    }

    public class Item
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public double Price { get; set; }
    }

    // Configuration class for reporting options
    public class ReportOptionsConfig
    {
        public bool InlineErrorMessages { get; set; }
        public bool UseReflectionOptimization { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "template.docx");
            string reportPath = Path.Combine(outputDir, "report.docx");
            string configPath = Path.Combine(outputDir, "reportOptions.json");

            // 1. Create JSON configuration file
            var config = new ReportOptionsConfig
            {
                InlineErrorMessages = true,
                UseReflectionOptimization = true
            };
            File.WriteAllText(configPath, JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true }));

            // 2. Create template document programmatically
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Title
            builder.Writeln("Order Report");
            builder.Writeln();

            // Customer name tag
            builder.Writeln("Customer: <<[order.CustomerName]>>");
            builder.Writeln();

            // Table header inside foreach
            builder.Writeln("<<foreach [item in order.Items]>>");
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Product");
            builder.InsertCell();
            builder.Writeln("Price");
            builder.EndRow();

            // Table row
            builder.InsertCell();
            builder.Writeln("<<[item.Index]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Price]>>");
            builder.EndRow();

            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save template
            templateDoc.Save(templatePath);

            // 3. Load template document
            var doc = new Document(templatePath);

            // 4. Load configuration
            var configJson = File.ReadAllText(configPath);
            var optionsConfig = JsonSerializer.Deserialize<ReportOptionsConfig>(configJson) ?? new ReportOptionsConfig();

            // 5. Configure ReportingEngine
            if (optionsConfig.UseReflectionOptimization)
                ReportingEngine.UseReflectionOptimization = true;

            var engine = new ReportingEngine();

            // Set engine options based on config
            ReportBuildOptions buildOptions = ReportBuildOptions.None;
            if (optionsConfig.InlineErrorMessages)
                buildOptions |= ReportBuildOptions.InlineErrorMessages;
            engine.Options = buildOptions;

            // 6. Prepare data model
            var order = new Order();

            // 7. Build report
            engine.BuildReport(doc, order, "order");

            // 8. Save generated report
            doc.Save(reportPath, SaveFormat.Docx);
        }
    }
}
