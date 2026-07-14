using System;
using System.IO;
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

        // Create a simple data model.
        var model = new Model { Name = "Aspose" };

        // Build a template document with a LINQ Reporting tag.
        var builder = new DocumentBuilder();
        builder.Writeln("<<[model.Name]>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // First build the report – this must succeed.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(reportPath);

        // Attempt to modify restricted types after a report has been built.
        try
        {
            // This call must fail because restricted types cannot be changed after BuildReport.
            ReportingEngine.SetRestrictedTypes(typeof(string));
            Console.WriteLine("SetRestrictedTypes succeeded unexpectedly.");
        }
        catch (InvalidOperationException ex)
        {
            // Expected path – the engine should throw an InvalidOperationException.
            Console.WriteLine($"Expected exception caught: {ex.Message}");
        }
        catch (ArgumentException ex)
        {
            // In some versions ArgumentException may be thrown; handle it as well.
            Console.WriteLine($"Expected exception caught (ArgumentException): {ex.Message}");
        }
    }
}

// Simple public data model used by the template.
public class Model
{
    public string Name { get; set; } = string.Empty;
}
