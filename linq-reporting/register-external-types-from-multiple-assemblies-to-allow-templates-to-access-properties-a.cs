using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json.Linq;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a template document with LINQ Reporting tags.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report generated on: <<[model.CurrentDate.ToString(\"yyyy-MM-dd\")]>>");
        builder.Writeln("External info name: <<[model.ExternalInfo.Name]>>");
        builder.Writeln("External info value: <<[model.ExternalInfo.Value]>>");
        builder.Writeln("JSON data (sample): <<[model.JsonData]>>");

        // Save the template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // 2. Load the template for reporting.
        Document doc = new Document(templatePath);

        // 3. Prepare the data model.
        var model = new ReportModel
        {
            CurrentDate = DateTime.Now,
            ExternalInfo = new ExternalInfo
            {
                Name = "SampleExternal",
                Value = 42
            },
            JsonData = new JObject
            {
                ["key"] = "value",
                ["number"] = 123
            }.ToString()
        };

        // 4. Register external types from different assemblies.
        ReportingEngine engine = new ReportingEngine();
        // Type from this assembly.
        engine.KnownTypes.Add(typeof(ExternalInfo));
        // Type from System (mscorlib) assembly.
        engine.KnownTypes.Add(typeof(DateTime));
        // Type from Newtonsoft.Json assembly.
        engine.KnownTypes.Add(typeof(JObject));

        // 5. Build the report.
        engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Data model exposed to the template.
public class ReportModel
{
    public DateTime CurrentDate { get; set; } = DateTime.MinValue;
    public ExternalInfo ExternalInfo { get; set; } = new();
    public string JsonData { get; set; } = string.Empty;
}

// Simulated external class from another assembly/namespace.
public class ExternalInfo
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
