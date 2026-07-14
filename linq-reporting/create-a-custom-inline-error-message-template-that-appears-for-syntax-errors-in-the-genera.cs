using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class InlineErrorMessageExample
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "John Doe";
    }

    public static void Main()
    {
        // Paths for the temporary template and the final report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "report.docx");

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Correct tag – will be replaced with the model's Name value.
        builder.Writeln("Customer: <<[model.Name]>>");

        // Intentional syntax error: unknown switch "-unknown". This will trigger an inline error message.
        builder.Writeln("This line contains a syntax error: <<[model.Name] -unknown>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report using the model as the data source.
        ReportModel model = new ReportModel();
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // The success flag will be false because of the syntax error,
        // but the document will contain the inline error messages.
        Console.WriteLine($"Report build successful: {success}");

        // -------------------------------------------------
        // 3. Save the generated report.
        // -------------------------------------------------
        loadedTemplate.Save(reportPath);
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
