using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for files used in the example.
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";
        const string configPath = "reportOptions.json";

        // -----------------------------------------------------------------
        // 1. Create a simple template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer: <<[model.CustomerName]>>");
        builder.Writeln("Order Id: <<[model.OrderId]>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Create a JSON configuration file that defines ReportingEngine options.
        // -----------------------------------------------------------------
        var sampleConfig = new ReportingConfig
        {
            Options = new List<string>
            {
                "RemoveEmptyParagraphs",
                "AllowMissingMembers"
            }
        };
        string jsonConfig = JsonConvert.SerializeObject(sampleConfig, Formatting.Indented);
        File.WriteAllText(configPath, jsonConfig);

        // -----------------------------------------------------------------
        // 3. Load the configuration at runtime.
        // -----------------------------------------------------------------
        string loadedJson = File.ReadAllText(configPath);
        ReportingConfig config = JsonConvert.DeserializeObject<ReportingConfig>(loadedJson)!;

        // Convert string option names to the corresponding enum flags.
        ReportBuildOptions engineOptions = ReportBuildOptions.None;
        foreach (string optName in config.Options)
        {
            if (Enum.TryParse(typeof(ReportBuildOptions), optName, out var parsed))
                engineOptions |= (ReportBuildOptions)parsed;
        }

        // -----------------------------------------------------------------
        // 4. Prepare the data model for the report.
        // -----------------------------------------------------------------
        var model = new OrderModel
        {
            CustomerName = "John Doe",
            OrderId = 12345
        };

        // -----------------------------------------------------------------
        // 5. Load the template document, configure the engine, and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = engineOptions
        };
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Configuration class that mirrors the JSON structure.
// ---------------------------------------------------------------------
public class ReportingConfig
{
    public List<string> Options { get; set; } = new();
}

// ---------------------------------------------------------------------
// Simple data model used by the template.
// ---------------------------------------------------------------------
public class OrderModel
{
    public string CustomerName { get; set; } = string.Empty;
    public int OrderId { get; set; }
}
