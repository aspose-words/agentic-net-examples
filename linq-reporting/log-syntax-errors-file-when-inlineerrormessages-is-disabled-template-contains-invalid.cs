using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Paths for the template, the generated document and the log file.
        string templatePath = "Template.docx";
        string outputPath = "Result.docx";
        string logPath = "SyntaxErrors.log";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Create a ReportingEngine without the InlineErrorMessages flag.
        // AllowMissingMembers is set only as an example; InlineErrorMessages is omitted.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };

        // Dummy data source – replace with real data as needed.
        var dataSource = new { };

        try
        {
            // Attempt to build the report. If the template contains invalid
            // expressions, an exception will be thrown because InlineErrorMessages is disabled.
            engine.BuildReport(doc, dataSource);
        }
        catch (Exception ex)
        {
            // Log the syntax error details to a file.
            File.AppendAllText(logPath,
                $"{DateTime.Now:u} - Syntax error: {ex.Message}{Environment.NewLine}");
        }

        // Save the (possibly unchanged) document.
        doc.Save(outputPath);
    }
}
