using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags, including an intentional syntax error
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Correct tag
        builder.Writeln("Customer: <<[model.CustomerName]>>");

        // Incorrect tag (missing closing bracket) to trigger an error
        builder.Writeln("Bad tag: <<[model.Missing>>");

        // Placeholder for inline error messages
        builder.Writeln("Error message: <<error>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            CustomerName = "John Doe"
        };

        // Configure the reporting engine to inline error messages
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report; the returned bool indicates success (true) or failure (false)
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(outputDir, "output.docx");
        reportDoc.Save(outputPath);

        // Output the result status
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// Public data model aligned with the template root name "model"
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
}
