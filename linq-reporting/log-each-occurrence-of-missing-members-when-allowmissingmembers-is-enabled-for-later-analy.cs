using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Existing member used in the template.
    public string Name { get; set; } = "John Doe";

    // Intentionally missing member to trigger logging.
    // public string MissingProperty { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Paths for template, output document, and log file.
        string templatePath = "template.docx";
        string outputPath = "output.docx";
        string logPath = "missing_members.log";

        // -----------------------------------------------------------------
        // 1. Create a simple template with a valid tag and a missing tag.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Hello, <<[model.Name]>>!");
        builder.Writeln("This will try to use a missing member: <<[model.MissingProperty]>>.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model (does NOT contain MissingProperty).
        // -----------------------------------------------------------------
        Model model = new();

        // -----------------------------------------------------------------
        // 4. Configure the ReportingEngine to allow missing members.
        // -----------------------------------------------------------------
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Retrieve and log each occurrence of missing members.
        // -----------------------------------------------------------------
        // Use reflection to obtain the MissingMemberLog property if it exists.
        IList<string> missingMembers = new List<string>();
        PropertyInfo? logProp = typeof(ReportingEngine).GetProperty("MissingMemberLog", BindingFlags.Instance | BindingFlags.Public);
        if (logProp != null && typeof(IList<string>).IsAssignableFrom(logProp.PropertyType))
        {
            missingMembers = (IList<string>?)logProp.GetValue(engine) ?? new List<string>();
        }

        // Ensure the log file directory exists.
        string logDirectory = Path.GetDirectoryName(Path.GetFullPath(logPath)) ?? "";
        if (!Directory.Exists(logDirectory))
        {
            Directory.CreateDirectory(logDirectory);
        }

        using (StreamWriter logWriter = new(logPath, false))
        {
            if (missingMembers.Count == 0)
            {
                logWriter.WriteLine("No missing members were encountered.");
                Console.WriteLine("No missing members were encountered.");
            }
            else
            {
                logWriter.WriteLine("Missing members encountered during report generation:");
                Console.WriteLine("Missing members encountered during report generation:");
                foreach (string entry in missingMembers)
                {
                    logWriter.WriteLine(entry);
                    Console.WriteLine(entry);
                }
            }
        }

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        // Ensure the output directory exists.
        string outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? "";
        if (!Directory.Exists(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        reportDoc.Save(outputPath);
    }
}
