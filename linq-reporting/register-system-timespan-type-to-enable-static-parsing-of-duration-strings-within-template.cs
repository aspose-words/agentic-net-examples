using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data model
        var model = new ReportModel
        {
            Title = "Project Schedule",
            Duration = TimeSpan.FromHours(5.5) // 5 hours 30 minutes
        };

        // Create a template document programmatically
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Report: <<[model.Title]>>");
        builder.Writeln("Total Duration: <<[model.Duration]>>");
        doc.Save(templatePath);

        // Load the template (simulating a separate load step)
        var template = new Document(templatePath);

        // Register System.TimeSpan type – not required for this simple example,
        // but shown here for completeness if a custom parser were needed.
        // ReportingEngine.RegisterCustomType(typeof(TimeSpan), value =>
        //     TimeSpan.TryParse(value, out var ts) ? ts : TimeSpan.Zero);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report
        var outputPath = "output.docx";
        template.Save(outputPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model used by the template
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public TimeSpan Duration { get; set; }
}
