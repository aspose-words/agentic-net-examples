using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create a simple template document containing a LINQ Reporting tag.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!"); // Tag will be replaced by the model's Name property.
        templateDoc.Save(templatePath); // Save the template to disk.

        // -------------------------------------------------
        // Load the template document for reporting.
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the data model.
        // -------------------------------------------------
        ReportModel model = new ReportModel { Name = "World" };

        // -------------------------------------------------
        // Disable reflection optimization temporarily within a scoped block.
        // -------------------------------------------------
        bool originalOptimization = ReportingEngine.UseReflectionOptimization;
        try
        {
            ReportingEngine.UseReflectionOptimization = false; // Turn off optimization.

            ReportingEngine engine = new ReportingEngine();
            // Build the report; the root object name must match the tag's reference ("model").
            engine.BuildReport(doc, model, "model");
        }
        finally
        {
            // Restore the original setting regardless of success/failure.
            ReportingEngine.UseReflectionOptimization = originalOptimization;
        }

        // -------------------------------------------------
        // Save the generated report.
        // -------------------------------------------------
        doc.Save(outputPath);
    }

    // Simple public data model class required by the template.
    public class ReportModel
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;
    }
}
