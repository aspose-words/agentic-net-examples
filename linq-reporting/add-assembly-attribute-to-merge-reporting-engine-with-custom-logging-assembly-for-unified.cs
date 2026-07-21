using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

// Assembly attribute to associate the reporting engine with a custom logger.
[assembly: ReportingEngineLogging(typeof(CustomLogger))]

public class ReportingEngineLoggingAttribute : Attribute
{
    public Type LoggerType { get; }

    public ReportingEngineLoggingAttribute(Type loggerType)
    {
        LoggerType = loggerType;
    }
}

public static class CustomLogger
{
    public static void Log(string message) => Console.WriteLine($"[CustomLog] {message}");
}

public class Model
{
    public string Name { get; set; } = "John Doe";
    // Intentionally missing property to trigger an inline error.
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new Model();

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<[model.Name]>>!");
        builder.Writeln("This will cause an error: <<[model.Unknown]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Retrieve the custom logger from the assembly attribute.
        var attr = (ReportingEngineLoggingAttribute?)Attribute.GetCustomAttribute(
            Assembly.GetExecutingAssembly(),
            typeof(ReportingEngineLoggingAttribute));

        // Log the result using the custom logger if available.
        if (attr?.LoggerType == typeof(CustomLogger))
        {
            if (success)
                CustomLogger.Log("Report generated successfully.");
            else
                CustomLogger.Log("Report generation failed. Inline error messages were inserted.");
        }

        // Save the generated report.
        var outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
