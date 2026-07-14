using System;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

// Assembly attribute that specifies the custom logger type.
// The attribute class is defined in the global namespace so it can be referenced here.
[assembly: LoggingAssemblyAttribute(typeof(CustomLogger))]

// Custom attribute to hold the logger type.
[AttributeUsage(AttributeTargets.Assembly, Inherited = false)]
public sealed class LoggingAssemblyAttribute : Attribute
{
    public Type LoggerType { get; }

    public LoggingAssemblyAttribute(Type loggerType) => LoggerType = loggerType;
}

// Simple logger that writes messages to the console.
public class CustomLogger
{
    public void Log(string message) => Console.WriteLine($"[CustomLogger] {message}");
}

namespace AsposeWordsLinqReportingExample
{
    // Sample data model.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
    }

    public static class Program
    {
        public static void Main()
        {
            // Resolve the logger from the assembly attribute.
            var loggerAttr = Assembly.GetExecutingAssembly()
                                     .GetCustomAttribute<LoggingAssemblyAttribute>();
            var logger = loggerAttr != null
                ? Activator.CreateInstance(loggerAttr.LoggerType) as CustomLogger
                : null;

            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document with a LINQ Reporting tag that will
            //    intentionally reference a missing member to trigger an error.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Customer Name: <<[person.Name]>>");
            builder.Writeln("Missing field (will cause error): <<[person.MissingProperty]>>");
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back from disk (required before building the report).
            // -----------------------------------------------------------------
            var loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            // Enable inline error messages so that BuildReport returns a success flag.
            engine.Options = ReportBuildOptions.InlineErrorMessages;
            // Provide a friendly message for missing members.
            engine.MissingMemberMessage = "N/A";

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            var data = new Person();
            bool success = engine.BuildReport(loadedTemplate, data, "person");

            // -----------------------------------------------------------------
            // 5. Unified error handling using the custom logger.
            // -----------------------------------------------------------------
            if (!success && logger != null)
            {
                logger.Log("Report generation completed with errors. See the document for inline messages.");
            }
            else if (logger != null)
            {
                logger.Log("Report generated successfully.");
            }

            // -----------------------------------------------------------------
            // 6. Save the final report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
