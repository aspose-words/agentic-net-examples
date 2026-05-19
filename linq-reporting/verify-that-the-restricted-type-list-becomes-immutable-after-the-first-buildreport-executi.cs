using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Model
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "Sample Name";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            // Insert a LINQ Reporting tag that references the model's Name property.
            builder.Writeln("<<[model.Name]>>");
            string templatePath = Path.Combine(outputDir, "template.docx");
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (demonstrates load rule usage).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Set restricted types BEFORE the first BuildReport call.
            // -----------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(string));

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            Model data = new Model();
            engine.BuildReport(loadedTemplate, data, "model");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, "report.docx");
            loadedTemplate.Save(reportPath);

            // -----------------------------------------------------------------
            // 5. Verify that the restricted type list is now immutable.
            // -----------------------------------------------------------------
            try
            {
                // Attempt to modify the restricted types after the first BuildReport.
                ReportingEngine.SetRestrictedTypes(typeof(int));
                Console.WriteLine("ERROR: Restricted types were modified after BuildReport (unexpected).");
            }
            catch (ArgumentException ex)
            {
                // Expected outcome – the list is immutable.
                Console.WriteLine("Expected exception caught: " + ex.Message);
            }
            catch (Exception ex)
            {
                // Any other exception type is also unexpected but reported.
                Console.WriteLine("Unexpected exception type: " + ex.GetType().Name + " - " + ex.Message);
            }

            // Indicate completion.
            Console.WriteLine("Example completed. Files written to: " + outputDir);
        }
    }
}
