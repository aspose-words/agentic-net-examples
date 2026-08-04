using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used in the template.
    public class Model
    {
        public string Name { get; set; } = "Aspose";
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // 0. Set restricted types BEFORE any Aspose.Words or ReportingEngine usage.
        // -----------------------------------------------------------------
        // This must be done at application startup to avoid the engine
        // marking the restricted‑type list as immutable.
        ReportingEngine.SetRestrictedTypes(typeof(System.Type));

        // Paths for the temporary template and output documents.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Insert a LINQ Reporting tag that references the model's Name property.
        builder.Writeln("Hello, <<[model.Name]>>!");
        // Save the template so it can be loaded later (required by the lifecycle rule).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the saved template document.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Build the report for the first time.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        Model model = new Model();
        // The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 4. Attempt to modify the restricted type list AFTER the first BuildReport.
        //    This should throw an exception because the list becomes immutable.
        // -----------------------------------------------------------------
        try
        {
            // Trying to set another restricted type should fail.
            ReportingEngine.SetRestrictedTypes(typeof(System.IO.FileInfo));
            Console.WriteLine("Restricted types were modified after BuildReport (unexpected).");
        }
        catch (InvalidOperationException ex)
        {
            // Expected outcome: the list is immutable.
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }

        // Clean up temporary files (optional).
        // File.Delete(templatePath);
        // File.Delete(outputPath);
    }
}
