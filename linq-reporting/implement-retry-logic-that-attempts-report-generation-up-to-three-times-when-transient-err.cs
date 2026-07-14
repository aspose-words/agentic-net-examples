using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = "Sample Report";
    public List<string> Items { get; set; } = new() { "Item 1", "Item 2", "Item 3" };
}

public class Program
{
    private const string TemplatePath = "Template.docx";
    private const string OutputPath = "Report.docx";

    public static void Main()
    {
        // Ensure the template exists.
        CreateTemplate();

        // Load the template document.
        var doc = new Document(TemplatePath);

        // Prepare the data source.
        var model = new ReportModel();

        // Initialize the reporting engine.
        var engine = new ReportingEngine();

        // Attempt to build the report with retry logic.
        const int maxAttempts = 3;
        int attempt = 0;
        bool success = false;

        while (attempt < maxAttempts && !success)
        {
            try
            {
                // Build the report. The root object name is "model".
                engine.BuildReport(doc, model, "model");
                success = true;
            }
            catch (Exception ex) when (IsTransient(ex))
            {
                attempt++;
                if (attempt >= maxAttempts)
                    throw new InvalidOperationException($"Report generation failed after {maxAttempts} attempts.", ex);
                // Optionally, wait before retrying (omitted for brevity).
            }
        }

        // Save the generated report.
        doc.Save(OutputPath);
    }

    // Determines whether an exception is transient.
    private static bool IsTransient(Exception ex)
    {
        // For demonstration, treat all exceptions as transient.
        // In real scenarios, inspect the exception type/message.
        return true;
    }

    // Creates a simple Word template with LINQ Reporting tags.
    private static void CreateTemplate()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title placeholder.
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln();

        // Items list using foreach.
        builder.Writeln("Items:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(TemplatePath);
    }
}
