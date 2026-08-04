using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple data model.
        ReportModel model = new ReportModel
        {
            Name = "John Doe"
        };

        // -----------------------------------------------------------------
        // Step 1: Create a template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Correct tag – will be replaced with the model's Name.
        builder.Writeln("Customer: <<[model.Name]>>");

        // Incorrect tag – the property Age does not exist in ReportModel.
        // This will generate a syntax error.
        builder.Writeln("Age: <<[model.Age]>>");

        // Placeholder where inline error messages will be inserted.
        builder.Writeln("<<error>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();
        // Enable inline error messages.
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // BuildReport returns a flag indicating whether parsing succeeded.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Output the result to the console.
        Console.WriteLine($"Report generation success flag: {success}");
        Console.WriteLine($"Template file: {Path.GetFullPath(templatePath)}");
        Console.WriteLine($"Report file:   {Path.GetFullPath(outputPath)}");
    }
}

// Simple data model used by the template.
public class ReportModel
{
    // The only property that actually exists.
    public string Name { get; set; } = string.Empty;
}
