using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Register code page provider for legacy encodings (required by Aspose.Words in some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the files used in the example.
        const string configPath = "reportOptions.json";
        const string templatePath = "template.docx";
        const string outputPath = "ReportOutput.docx";

        // -----------------------------------------------------------------
        // 1. Create a simple JSON configuration file that specifies reporting options.
        // -----------------------------------------------------------------
        var sampleConfig = new ReportingConfig
        {
            Options = new List<string> { "RemoveEmptyParagraphs", "AllowMissingMembers" }
        };
        File.WriteAllText(configPath, JsonConvert.SerializeObject(sampleConfig, Formatting.Indented));

        // -----------------------------------------------------------------
        // 2. Load the configuration at runtime.
        // -----------------------------------------------------------------
        ReportingConfig config = JsonConvert.DeserializeObject<ReportingConfig>(File.ReadAllText(configPath));

        // Convert the list of option names into a combined ReportBuildOptions flag value.
        ReportBuildOptions engineOptions = ReportBuildOptions.None;
        if (config?.Options != null)
        {
            foreach (string optName in config.Options)
            {
                if (Enum.TryParse(optName, out ReportBuildOptions parsed))
                    engineOptions |= parsed;
            }
        }

        // -----------------------------------------------------------------
        // 3. Create a Word template programmatically with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Customer Report");
        builder.Writeln("----------------");
        builder.Writeln("<<foreach [c in Customers]>>");
        builder.Writeln("Name: <<[c.Name]>>");
        builder.Writeln("Age : <<[c.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required by the lifecycle rule).
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template back (simulating a real-world scenario).
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 5. Prepare the data source.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Customers = new List<Customer>
            {
                new Customer { Name = "Alice", Age = 30 },
                new Customer { Name = "Bob",   Age = 45 },
                new Customer { Name = "Carol", Age = 27 }
            }
        };

        // -----------------------------------------------------------------
        // 6. Configure the ReportingEngine with the loaded options.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = engineOptions
        };

        // Optional: customize the message for missing members when the option is enabled.
        if ((engineOptions & ReportBuildOptions.AllowMissingMembers) != 0)
            engine.MissingMemberMessage = "N/A";

        // -----------------------------------------------------------------
        // 7. Build the report.
        // -----------------------------------------------------------------
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 8. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }

    // Class representing the JSON configuration structure.
    private class ReportingConfig
    {
        public List<string> Options { get; set; } = new();
    }

    // Root data model used by the template.
    public class ReportModel
    {
        public List<Customer> Customers { get; set; } = new();
    }

    // Simple data entity.
    public class Customer
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }
}
