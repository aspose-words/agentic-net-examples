using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model – only Name is defined.
    public class Model
    {
        public string Name { get; set; } = "John Doe";
        // Age is intentionally omitted to trigger a missing‑member warning.
    }

    public static void Main()
    {
        // Paths for the temporary template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath   = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Normal field – will be filled correctly.
        builder.Writeln("Name: <<[model.Name]>>");

        // Missing field – there is no Age property in Model.
        // This will cause a warning/error during report generation.
        builder.Writeln("Age: <<[model.Age]>>");

        // The <<error>> tag will be replaced with the inline error message
        // when the ReportingEngine is configured with InlineErrorMessages.
        builder.Writeln("<<error>>");

        // Save the template to disk (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template back before building the report.
        // -------------------------------------------------
        var doc = new Document(templatePath);

        // -------------------------------------------------
        // 3. Prepare the data source.
        // -------------------------------------------------
        var model = new Model();

        // -------------------------------------------------
        // 4. Configure and run the ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine
        {
            // InlineErrorMessages makes the engine insert error messages
            // directly into the document where parsing problems occur.
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns a bool indicating success when InlineErrorMessages is set.
        bool success = engine.BuildReport(doc, model, "model");

        // -------------------------------------------------
        // 5. Save the generated report.
        // -------------------------------------------------
        doc.Save(reportPath);

        // Output the result to the console.
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}.");
        Console.WriteLine($"Template:  {Path.GetFullPath(templatePath)}");
        Console.WriteLine($"Report:    {Path.GetFullPath(reportPath)}");
    }
}
