using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -------------------------------------------------
        // 1. Create the template document with LINQ tags
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple title
        builder.Writeln("LINQ Reporting – Inline Error Demo");
        builder.Writeln();

        // Conditional block that references a missing property (will cause an error)
        builder.Writeln("<<if [model.Missing]>>");
        builder.Writeln("This text is inside a failing condition.");
        // <<error>> tag will be replaced with the evaluation error message
        builder.Writeln("<<error>>");
        builder.Writeln("<</if>>");
        builder.Writeln();

        // Conditional block that evaluates correctly
        builder.Writeln("<<if [model.Value > 0]>>");
        builder.Writeln("The supplied value is: <<[model.Value]>>");
        builder.Writeln("<</if>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model (intentionally missing the 'Missing' property)
        var model = new ReportModel { Value = 42 };

        // Configure the reporting engine to inline error messages
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report; the root object name is "model"
        bool success = engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 3. Save the generated report
        // -------------------------------------------------
        reportDoc.Save(outputPath);

        // Optional: indicate success/failure in console (no user interaction required)
        Console.WriteLine(success
            ? "Report generated successfully."
            : "Report generated with errors (see inline messages).");
    }

    // Data model used by the template
    public class ReportModel
    {
        // Existing property referenced by a valid condition
        public int Value { get; set; }

        // Note: No 'Missing' property is defined on purpose to trigger an error
    }
}
