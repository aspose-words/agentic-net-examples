using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

// Assembly attribute that specifies the custom logger type for the reporting engine.
[assembly: ReportingEngineLoggingAttribute(typeof(CustomLogger))]

/// <summary>
/// Assembly‑level attribute used to indicate which logger type should be instantiated
/// for unified error handling in the reporting engine.
/// </summary>
public sealed class ReportingEngineLoggingAttribute : Attribute
{
    public Type LoggerType { get; }

    public ReportingEngineLoggingAttribute(Type loggerType) => LoggerType = loggerType;
}

/// <summary>
/// Simple logger interface used by the example.
/// </summary>
public interface ICustomLogger
{
    void Log(string message);
}

/// <summary>
/// Concrete logger that writes messages to the console.
/// </summary>
public class CustomLogger : ICustomLogger
{
    public void Log(string message) => Console.WriteLine($"[LOG] {message}");
}

namespace AsposeWordsLinqReportingExample
{
    // Sample data model used by the template.
    public class Model
    {
        public string Name { get; set; } = "Sample Name";
    }

    public static class Program
    {
        public static void Main()
        {
            // Resolve the logger from the assembly attribute.
            ICustomLogger? logger = ResolveLogger();

            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // Step 1: Create a template document with a LINQ Reporting tag that will cause an error
            // (referencing a missing property) to demonstrate unified error handling.
            CreateTemplate(templatePath, logger);

            // Step 2: Load the template document.
            Document template = new Document(templatePath);

            // Step 3: Prepare the data source.
            Model model = new Model();

            // Step 4: Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;
            engine.MissingMemberMessage = "Property not found";

            // Step 5: Build the report.
            bool success = engine.BuildReport(template, model, "model");

            // Step 6: Log the result.
            if (success)
                logger?.Log("Report generated successfully.");
            else
                logger?.Log("Report generation failed due to template errors.");

            // Step 7: Save the generated report.
            template.Save(outputPath);
        }

        // Retrieves the logger instance defined by the assembly attribute.
        private static ICustomLogger? ResolveLogger()
        {
            var attr = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<ReportingEngineLoggingAttribute>();

            if (attr != null && typeof(ICustomLogger).IsAssignableFrom(attr.LoggerType))
                return (ICustomLogger?)Activator.CreateInstance(attr.LoggerType);

            return null;
        }

        // Creates a simple Word template containing a LINQ Reporting tag.
        private static void CreateTemplate(string path, ICustomLogger? logger)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Normal text.
            builder.Writeln("Report for: <<[model.Name]>>");

            // This tag references a non‑existent property and will trigger an inline error message.
            builder.Writeln("Missing data: <<[model.NonExistentProperty]>>");

            // Save the template.
            doc.Save(path);
            logger?.Log($"Template created at '{Path.GetFullPath(path)}'.");
        }
    }
}
