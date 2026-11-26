using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // Paths for the template and the final report.
        string templatePath = Path.Combine(workDir, "Template.docx");
        string reportPath = Path.Combine(workDir, "Report.docx");

        // Path to the optional include document (do NOT create it – it may be missing).
        string optionalIncludePath = Path.Combine(workDir, "OptionalInclude.docx");

        // -----------------------------------------------------------------
        // Build the template document.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Report start");

        // Conditional inclusion – the document will be inserted only if the file exists.
        builder.Writeln("<<if [model.IncludeFileExists]>>");
        builder.Writeln("<<doc [model.IncludePath]>>");
        builder.Writeln("<</if>>");

        builder.Writeln("Report end");

        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Prepare the model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            IncludePath = optionalIncludePath,
            IncludeFileExists = File.Exists(optionalIncludePath)
        };

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Configure the reporting engine (no special options needed for this scenario).
        ReportingEngine engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);

        // Output the result.
        Console.WriteLine($"Report generated successfully: {success}");
        Console.WriteLine($"Template path: {templatePath}");
        Console.WriteLine($"Report path:   {reportPath}");
    }

    // Model class used by the reporting engine.
    public class ReportModel
    {
        // Path to the optional include document.
        public string IncludePath { get; set; } = string.Empty;

        // Indicates whether the include file exists; if false the <<doc>> tag is skipped.
        public bool IncludeFileExists { get; set; }
    }
}
