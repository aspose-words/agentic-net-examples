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
        // 1. Build the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Tag that references a missing member – will be treated as null.
        builder.Writeln("<<[missingObject.First().id]>>");

        // Foreach loop over a missing collection – also treated as null/empty.
        builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

        // Save the template to disk (required before BuildReport according to rules).
        string templatePath = System.IO.Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and configure the ReportingEngine.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine
        {
            // Treat missing members as null literals.
            Options = ReportBuildOptions.AllowMissingMembers,
            // Optional: custom text to display for a missing plain member reference.
            MissingMemberMessage = "Missed"
        };

        // Use an empty DataSet as the data source – it contains no members.
        DataSet emptyData = new DataSet();

        // Build the report. The empty string for dataSourceName means we do not
        // need to reference the data source object itself in the template.
        bool success = engine.BuildReport(loadedTemplate, emptyData, "");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        string reportPath = System.IO.Path.Combine(outputDir, "Report.docx");
        loadedTemplate.Save(reportPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
