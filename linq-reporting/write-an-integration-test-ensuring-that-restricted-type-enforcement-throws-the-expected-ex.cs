using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingTest
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        // Exposes a System.Type instance; we will restrict System.Type later.
        public Type TypeVar { get; set; } = typeof(string);
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document with a tag that accesses a member of System.Type.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            CreateTemplate(templatePath);

            // 2. Load the template document.
            Document doc = new Document(templatePath);

            // 3. Set restricted types before any report generation.
            // Restrict System.Type so that its members cannot be accessed from the template.
            ReportingEngine.SetRestrictedTypes(typeof(Type));

            // 4. Prepare the model instance.
            ReportModel model = new ReportModel();

            // 5. Build the report and verify that an exception is thrown.
            try
            {
                ReportingEngine engine = new ReportingEngine();
                // The template uses the root name "model".
                engine.BuildReport(doc, model, "model");

                // If no exception is thrown, the test has failed.
                Console.WriteLine("Test FAILED: No exception was thrown when accessing a restricted type.");
            }
            catch (Exception ex)
            {
                // Expected path – an exception should be thrown because System.Type is restricted.
                Console.WriteLine("Test PASSED: Caught expected exception.");
                Console.WriteLine($"Exception type: {ex.GetType().FullName}");
                Console.WriteLine($"Message: {ex.Message}");
            }
        }

        // Creates a simple Word document containing a LINQ Reporting tag that accesses System.Type.FullName.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The tag attempts to read the FullName property of the TypeVar property.
            // Since System.Type will be restricted, this should cause an exception during BuildReport.
            builder.Writeln("<<[model.TypeVar.FullName]>>");

            // Save the template to disk.
            doc.Save(filePath);
        }
    }
}
