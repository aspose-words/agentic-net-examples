using System;
using System.Data;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string reportPath = Path.Combine(workDir, "Report.docx");
        string logPath = Path.Combine(workDir, "MissingMembersLog.txt");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains references to missing members.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Tag that tries to access a missing object.
        builder.Writeln("<<[missingObject.First().id]>>");

        // Foreach loop over a missing collection.
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Save the template to disk (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back and build the report with AllowMissingMembers.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missed"
        };

        // Use an empty DataSet as the data source – it does not contain the missing members.
        bool success = engine.BuildReport(doc, new DataSet(), "");

        // Save the generated report.
        doc.Save(reportPath);

        // -----------------------------------------------------------------
        // 3. Log each occurrence of the missing‑member placeholder.
        // -----------------------------------------------------------------
        // The engine replaces every missing member with the value of MissingMemberMessage.
        string placeholder = engine.MissingMemberMessage;
        string documentText = doc.GetText();

        // Find all lines that contain the placeholder.
        var linesWithMissing = documentText
            .Split(new[] { "\r\n", "\n" }, StringSplitOptions.None)
            .Select((line, index) => new { Line = line, Index = index + 1 })
            .Where(l => l.Line.Contains(placeholder))
            .ToList();

        // Write log to console.
        Console.WriteLine($"Missing member placeholder \"{placeholder}\" found {linesWithMissing.Count} time(s).");
        foreach (var entry in linesWithMissing)
        {
            Console.WriteLine($"Line {entry.Index}: {entry.Line.Trim()}");
        }

        // Write log to a file for later analysis.
        using (StreamWriter writer = new StreamWriter(logPath, false))
        {
            writer.WriteLine($"Report generated at: {DateTime.Now}");
            writer.WriteLine($"Template: {templatePath}");
            writer.WriteLine($"Report: {reportPath}");
            writer.WriteLine();
            writer.WriteLine($"Missing member placeholder \"{placeholder}\" occurrences: {linesWithMissing.Count}");
            foreach (var entry in linesWithMissing)
            {
                writer.WriteLine($"Line {entry.Index}: {entry.Line.Trim()}");
            }
        }

        // Indicate completion.
        Console.WriteLine("Processing complete. Log saved to: " + logPath);
    }
}
