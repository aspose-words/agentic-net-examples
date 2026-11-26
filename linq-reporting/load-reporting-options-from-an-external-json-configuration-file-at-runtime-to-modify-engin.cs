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
        // Register code page provider (required for Aspose.Words)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a simple JSON configuration file with reporting options
        string configPath = Path.Combine(outputDir, "reportConfig.json");
        File.WriteAllText(configPath, @"{ ""Options"": [ ""InlineErrorMessages"" ] }");

        // Create a simple Word template with LINQ Reporting tags
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report generated for: <<[model.Name]>>");
        builder.Writeln("Current date: <<[model.Date]>>");
        templateDoc.Save(templatePath);

        // Load the template document
        Document doc = new Document(templatePath);

        // Load reporting options from JSON configuration
        string json = File.ReadAllText(configPath);
        ReportConfig config = JsonConvert.DeserializeObject<ReportConfig>(json) ?? new ReportConfig();

        ReportBuildOptions options = 0;
        foreach (string opt in config.Options)
        {
            if (Enum.TryParse(opt, out ReportBuildOptions flag))
                options |= flag;
        }

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            Name = "Alice Johnson",
            Date = DateTime.Now.ToString("D")
        };

        // Configure and run the reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.Options = options;
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string resultPath = Path.Combine(outputDir, "report.docx");
        doc.Save(resultPath);
    }
}

// Configuration class matching the JSON structure
public class ReportConfig
{
    public List<string> Options { get; set; } = new();
}

// Simple data model used by the template
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public string Date { get; set; } = string.Empty;
}
