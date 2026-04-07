using System;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

[assembly: ReportingEngineLogging(typeof(CustomLogger))]

public sealed class ReportingEngineLoggingAttribute : Attribute
{
    public Type LoggerType { get; }

    public ReportingEngineLoggingAttribute(Type loggerType) => LoggerType = loggerType;
}

public static class CustomLogger
{
    public static void Log(string message) => Console.WriteLine($"[CustomLogger] {message}");
}

public class ReportModel
{
    public string Name { get; set; } = "World";
    // Intentionally missing property to trigger an error in the template.
}

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a simple template document with LINQ Reporting tags.
        var templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // 2. Load the template.
        var doc = new Document(templatePath);

        // 3. Prepare the data model.
        var model = new ReportModel();

        // 4. Configure the reporting engine to emit inline error messages.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // 5. Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // 6. If there were errors, log them using the custom logger obtained from the assembly attribute.
        if (!success)
        {
            var attr = Assembly.GetExecutingAssembly()
                .GetCustomAttribute<ReportingEngineLoggingAttribute>();

            if (attr?.LoggerType != null)
            {
                var logMethod = attr.LoggerType.GetMethod("Log", BindingFlags.Public | BindingFlags.Static);
                logMethod?.Invoke(null, new object[] { "Report generation completed with errors. Check the output document for inline messages." });
            }
        }

        // 7. Save the generated report.
        var outputPath = Path.Combine(outputDir, "output.docx");
        doc.Save(outputPath);

        // Inform the user (no interactive wait).
        Console.WriteLine($"Report saved to: {outputPath}");
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Hello <<[model.Name]>>!");
        builder.Writeln("This line will cause an error: <<[model.MissingProperty]>>");

        doc.Save(path);
    }
}
