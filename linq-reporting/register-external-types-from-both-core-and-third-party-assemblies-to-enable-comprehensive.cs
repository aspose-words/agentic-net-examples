using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json.Linq;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        var model = new ReportModel
        {
            Title = "Sample Report",
            JsonData = JObject.Parse("{\"Name\":\"John Doe\",\"Age\":30}")
        };
        // Populate convenient properties from JSON
        model.Name = model.JsonData["Name"]?.ToString() ?? string.Empty;
        model.Age = model.JsonData["Age"]?.ToObject<int>() ?? 0;

        // Create a template document with LINQ Reporting tags
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("Current Date: <<[model.CurrentDate]>>");
        builder.Writeln("Name from JSON: <<[model.Name]>>");
        builder.Writeln("Age from JSON: <<[model.Age]>>");

        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }

    public class ReportModel
    {
        public string Title { get; set; } = string.Empty;
        public JObject JsonData { get; set; } = new();
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public DateTime CurrentDate => DateTime.Now;
    }
}
