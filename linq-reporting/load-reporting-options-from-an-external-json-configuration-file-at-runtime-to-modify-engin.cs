using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingDemo
{
    // Model class used as the data source for the report.
    public class Order
    {
        public string CustomerName { get; set; } = "John Doe";
        public decimal Amount { get; set; } = 123.45m;
    }

    // Configuration class that mirrors the JSON file structure.
    public class ReportingOptionsConfig
    {
        public bool AllowMissingMembers { get; set; }
        public bool RemoveEmptyParagraphs { get; set; }
        public bool InlineErrorMessages { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings used by Aspose.Words).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Prepare sample JSON configuration file.
            // -----------------------------------------------------------------
            const string configPath = "reportingOptions.json";
            var sampleConfig = new ReportingOptionsConfig
            {
                AllowMissingMembers = true,
                RemoveEmptyParagraphs = true,
                InlineErrorMessages = false
            };
            File.WriteAllText(configPath, JsonConvert.SerializeObject(sampleConfig, Formatting.Indented));

            // -----------------------------------------------------------------
            // 2. Create a simple Word template with LINQ Reporting tags.
            // -----------------------------------------------------------------
            const string templatePath = "template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Customer: <<[model.CustomerName]>>");
            builder.Writeln("Amount: <<[model.Amount]>>");
            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 3. Load reporting options from the JSON configuration file.
            // -----------------------------------------------------------------
            var configJson = File.ReadAllText(configPath);
            var optionsConfig = JsonConvert.DeserializeObject<ReportingOptionsConfig>(configJson)!;

            // Build the combined ReportBuildOptions flag based on the configuration.
            ReportBuildOptions engineOptions = ReportBuildOptions.None;
            if (optionsConfig.AllowMissingMembers)
                engineOptions |= ReportBuildOptions.AllowMissingMembers;
            if (optionsConfig.RemoveEmptyParagraphs)
                engineOptions |= ReportBuildOptions.RemoveEmptyParagraphs;
            if (optionsConfig.InlineErrorMessages)
                engineOptions |= ReportBuildOptions.InlineErrorMessages;

            // -----------------------------------------------------------------
            // 4. Initialize the ReportingEngine with the loaded options.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine { Options = engineOptions };

            // -----------------------------------------------------------------
            // 5. Prepare the data source object.
            // -----------------------------------------------------------------
            var order = new Order();

            // -----------------------------------------------------------------
            // 6. Load the template document (demonstrating load after creation).
            // -----------------------------------------------------------------
            var docForReport = new Document(templatePath);

            // -----------------------------------------------------------------
            // 7. Build the report using the engine.
            // -----------------------------------------------------------------
            engine.BuildReport(docForReport, order, "model");

            // -----------------------------------------------------------------
            // 8. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "output/report.docx";
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
            docForReport.Save(outputPath);
        }
    }
}
