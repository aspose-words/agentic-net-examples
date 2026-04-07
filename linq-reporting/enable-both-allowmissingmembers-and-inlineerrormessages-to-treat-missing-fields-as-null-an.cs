using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the files.
        string outputDir = "Output";
        System.IO.Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Tag that references a missing object – will be treated as null.
        builder.Writeln("Missing object first ID: <<[missingObject.First().id]>>");

        // Foreach loop over a missing collection – each iteration will output nothing.
        builder.Writeln("Missing collection loop:");
        builder.Writeln("<<foreach [in missingObject]>>");
        builder.Writeln("  Item ID: <<[id]>>");
        builder.Writeln("<</foreach>>");

        // Tag with a syntax error (unsupported switch) – will be shown inline.
        builder.Writeln("Syntax error example: <<[missingObject.id] -unknown>>");

        // Save the template to disk.
        string templatePath = System.IO.Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back (required before building the report).
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages;
        engine.MissingMemberMessage = "Missed";

        // Use an empty DataSet as the data source – it contains no members.
        DataSet dataSource = new DataSet();

        // Build the report. The returned flag is meaningful because InlineErrorMessages is enabled.
        bool success = engine.BuildReport(loadedTemplate, dataSource, "");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = System.IO.Path.Combine(outputDir, "Report.docx");
        loadedTemplate.Save(reportPath);

        // Output simple information to the console (no user interaction required).
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
