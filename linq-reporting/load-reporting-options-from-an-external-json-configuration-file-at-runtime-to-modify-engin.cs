using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportingConfig
{
    public bool InlineErrorMessages { get; set; }
    public bool UseReflectionOptimization { get; set; }
}

public class Order
{
    public string CustomerName { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare JSON configuration file.
        string configPath = "reportConfig.json";
        string configJson = @"{
            ""InlineErrorMessages"": true,
            ""UseReflectionOptimization"": true
        }";
        File.WriteAllText(configPath, configJson);

        // Load configuration.
        ReportingConfig config = JsonConvert.DeserializeObject<ReportingConfig>(File.ReadAllText(configPath));

        // Create a simple template document with a LINQ Reporting tag.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer: <<[order.CustomerName]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Sample data model.
        Order order = new Order { CustomerName = "John Doe" };

        // Configure the reporting engine based on JSON options.
        ReportingEngine engine = new ReportingEngine();
        if (config.InlineErrorMessages)
            engine.Options = ReportBuildOptions.InlineErrorMessages;
        else
            engine.Options = ReportBuildOptions.None;

        if (config.UseReflectionOptimization)
            ReportingEngine.UseReflectionOptimization = true;

        // Build the report.
        bool success = engine.BuildReport(doc, order, "order");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to {outputPath}");
    }
}
