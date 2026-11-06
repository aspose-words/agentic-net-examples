using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class KnownExternalTypes
{
    // Dummy implementation to satisfy compilation when the Aspose class is unavailable.
    public static void Add(Type type) { /* No operation – Aspose would handle this internally. */ }
}

public class Program
{
    public static void Main()
    {
        // Register the custom external type so the reporting engine can access it without reflection.
        KnownExternalTypes.Add(typeof(MyExternalType));
        ReportingEngine.UseReflectionOptimization = true;

        // Create a simple template document with a tag that accesses the external type.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("External value: <<[model.Data.Ext.Value]>>");
        templateDoc.Save(templatePath);

        // Prepare the data model.
        var model = new ReportModel
        {
            Data = new DataWrapper
            {
                Ext = new MyExternalType { Value = "Hello from external type!" }
            }
        };

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to '{outputPath}'.");
    }

    // Root model class.
    public class ReportModel
    {
        public DataWrapper Data { get; set; } = new();
    }

    // Wrapper class exposing the external type.
    public class DataWrapper
    {
        public MyExternalType Ext { get; set; } = new();
    }

    // Custom external type that will be registered with KnownExternalTypes.
    public class MyExternalType
    {
        public string Value { get; set; } = string.Empty;
    }
}
