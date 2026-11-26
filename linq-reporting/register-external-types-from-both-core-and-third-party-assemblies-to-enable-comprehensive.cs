using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json.Linq;

public class Program
{
    public static void Main()
    {
        // Create output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build the template document with LINQ Reporting tags
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Value of Math.PI: <<[model.MathPI]>>");
        builder.Writeln("JSON value (JObject.Parse): <<[model.JsonValue]>>");

        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Prepare the model that supplies the data for the template
        ReportModel model = new ReportModel();

        // Load the template for report generation
        Document reportDoc = new Document(templatePath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }

    // Model class exposing the values used in the template
    public class ReportModel
    {
        public double MathPI { get; set; } = Math.PI;
        public int JsonValue { get; set; }

        public ReportModel()
        {
            JObject obj = JObject.Parse("{\"value\":123}");
            JsonValue = (int)obj["value"]!;
        }
    }
}
