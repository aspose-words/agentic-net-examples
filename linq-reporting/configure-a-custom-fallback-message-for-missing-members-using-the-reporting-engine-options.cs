using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Customer: <<[model.Name]>>");
        // This tag references a missing member; we will enable inline error messages.
        builder.Writeln("Missing: <<[model.MissingProp]>>");
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        var template = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        var model = new Model();

        // -----------------------------------------------------------------
        // 4. Configure the reporting engine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // Enable inline error messages so that missing members are shown in the output.
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // NOTE: In newer versions of Aspose.Words a custom fallback message can be set
        // via the ReportingEngine.FallbackMessage property. This property is not
        // available in the version used for this example, so we rely on the default
        // inline error message format.

        // -----------------------------------------------------------------
        // 5. Build the report.
        // -----------------------------------------------------------------
        bool success = engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // 6. Save the generated report.
        // -----------------------------------------------------------------
        template.Save(reportPath);

        // Output the result status.
        Console.WriteLine($"Report generation success: {success}");
    }
}

// ---------------------------------------------------------------------
// Data model used by the template.
// ---------------------------------------------------------------------
public class Model
{
    // Existing property referenced by the template.
    public string Name { get; set; } = "John Doe";

    // MissingProp is intentionally omitted to trigger the inline error message.
}
