using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

// Assembly attribute to specify the custom logger for the reporting engine.
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
    public static void Log(string message)
    {
        // Simple console logging; in real scenarios this could write to a file or logging framework.
        Console.WriteLine($"[CustomLogger] {message}");
    }
}

// Sample data model.
public class ReportModel
{
    public string Name { get; set; } = "World";
    // Intentionally omitted property to trigger an inline error.
    // public string Missing { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a temporary folder for files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!");
        // This tag references a non‑existent property to demonstrate error handling.
        builder.Writeln("Missing property: <<[model.Missing]>>");
        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Retrieve the custom logger from the assembly attribute.
        var loggingAttr = Assembly.GetExecutingAssembly()
            .GetCustomAttribute<ReportingEngineLoggingAttribute>();
        Action<string>? logAction = null;
        if (loggingAttr != null && typeof(CustomLogger).IsAssignableFrom(loggingAttr.LoggerType))
        {
            // Use reflection to obtain the static Log method.
            MethodInfo? logMethod = loggingAttr.LoggerType.GetMethod("Log", BindingFlags.Public | BindingFlags.Static);
            if (logMethod != null)
            {
                logAction = (msg) => logMethod.Invoke(null, new object[] { msg });
            }
        }

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Log the result using the custom logger if available.
        if (logAction != null)
        {
            if (success)
                logAction("Report built successfully.");
            else
                logAction("Report build failed; see inline error messages in the document.");
        }

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "Result.docx");
        doc.Save(resultPath);
    }
}
