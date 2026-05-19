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
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Tag that references an existing member.
        builder.Writeln("Customer Name: <<[model.ExistingProperty]>>");
        // Tag that references a missing member – this will cause an exception.
        builder.Writeln("Missing Member: <<[model.MissingProperty]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document for reporting.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        var model = new ReportModel();

        // -----------------------------------------------------------------
        // 4. Build the report without AllowMissingMembers.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine
        {
            // Do NOT enable AllowMissingMembers; default is None.
            Options = ReportBuildOptions.None
        };

        try
        {
            // This call is expected to throw because MissingProperty does not exist.
            engine.BuildReport(doc, model, "model");

            // If no exception occurs, save the generated report.
            doc.Save(reportPath);
            Console.WriteLine("Report generated successfully.");
        }
        catch (Exception ex)
        {
            // Expected path: an exception is thrown for the missing member.
            Console.WriteLine($"Exception caught as expected: {ex.Message}");
        }
    }

    // Data model used by the report. Contains only ExistingProperty.
    public class ReportModel
    {
        public string ExistingProperty { get; set; } = "John Doe";
        // Note: MissingProperty is intentionally omitted.
    }
}
