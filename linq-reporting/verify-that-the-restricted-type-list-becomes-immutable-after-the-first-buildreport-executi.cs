using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 0. Set restricted types BEFORE any Aspose.Words operations.
            // -----------------------------------------------------------------
            // This must be done at application startup, before any report is built.
            ReportingEngine.SetRestrictedTypes(typeof(string));

            // Paths for the template and the generated report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // Insert a simple LINQ Reporting tag that references the model.
            builder.Writeln("Hello <<[model.Name]>>!");
            // Save the template to disk (required by the rule set).
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Build the report using the model as the root data source.
            // -----------------------------------------------------------------
            var model = new Model(); // Name = "World"
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save(reportPath);

            // -----------------------------------------------------------------
            // 4. Verify that the restricted type list is now immutable.
            //    Attempting to modify it should throw an InvalidOperationException.
            // -----------------------------------------------------------------
            try
            {
                // This call must fail because BuildReport has already been executed.
                ReportingEngine.SetRestrictedTypes(typeof(int));
                Console.WriteLine("ERROR: Restricted types were modified after BuildReport.");
            }
            catch (InvalidOperationException ex)
            {
                // Expected path – the list is immutable after the first BuildReport execution.
                Console.WriteLine("Restricted types are immutable after the first BuildReport execution.");
                Console.WriteLine("Exception message: " + ex.Message);
            }
        }
    }
}
