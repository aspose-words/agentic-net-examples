using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template with a missing member reference
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Prepare data model (does NOT contain the missing member)
        ReportModel model = new ReportModel
        {
            ExistingProperty = "I exist"
        };

        // Prepare a list to collect missing member logs
        List<string> missingMemberLogs = new List<string>();

        // Detect missing members by inspecting the template tags before building the report
        DetectMissingMembers(doc, model, missingMemberLogs);

        // Configure the reporting engine
        ReportingEngine engine = new ReportingEngine();

        // Enable AllowMissingMembers option
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        // Write logs to a file
        string logPath = Path.Combine(outputDir, "missing.log");
        File.WriteAllLines(logPath, missingMemberLogs);

        // Output summary
        Console.WriteLine("Report generation completed.");
        Console.WriteLine($"Success: {success}");
        Console.WriteLine($"Report saved to: {reportPath}");
        Console.WriteLine($"Missing member log saved to: {logPath}");
        if (missingMemberLogs.Count > 0)
        {
            Console.WriteLine("Missing members encountered:");
            foreach (var line in missingMemberLogs)
                Console.WriteLine(line);
        }
        else
        {
            Console.WriteLine("No missing members were encountered.");
        }
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("=== LINQ Reporting Example ===");
        builder.Writeln("Existing property: <<[model.ExistingProperty]>>");
        // This tag references a non‑existent member and will trigger the missing‑member handling.
        builder.Writeln("Missing property: <<[model.NonExistentProperty]>>");

        doc.Save(path);
    }

    // Simple data model with only one existing property
    public class ReportModel
    {
        public string ExistingProperty { get; set; } = string.Empty;
    }

    private static void DetectMissingMembers(Document doc, ReportModel model, List<string> logs)
    {
        // Extract all LINQ Reporting expression tags from the document text
        string docText = doc.GetText();
        // Pattern matches tags like <<[model.PropertyName]>>
        Regex tagRegex = new Regex(@"<<\[\s*model\.([A-Za-z_][A-Za-z0-9_]*)\s*\]>>", RegexOptions.Compiled);
        MatchCollection matches = tagRegex.Matches(docText);

        Type modelType = typeof(ReportModel);
        foreach (Match match in matches)
        {
            string memberName = match.Groups[1].Value;
            // Check if the property exists on the model
            if (modelType.GetProperty(memberName) == null)
            {
                string log = $"Missing member detected in template: '{memberName}'.";
                logs.Add(log);
            }
        }
    }
}
