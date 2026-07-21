using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used as the root data source for the report.
    public class Person
    {
        // Initialize to avoid nullable warnings.
        public string Name { get; set; } = "World";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths in the current working directory.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");

            // -----------------------------------------------------------------
            // 1. Create a minimal Word template containing a LINQ Reporting tag.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // The tag references the data source named "model" and its Name property.
            builder.Writeln("Hello <<[model.Name]>>!");
            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real-world scenario).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Set restricted types BEFORE the first BuildReport call.
            // -----------------------------------------------------------------
            // Restrict access to the System.String type (any type can be used here).
            ReportingEngine.SetRestrictedTypes(typeof(string));

            // -----------------------------------------------------------------
            // 4. Build the report using the loaded template.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // BuildReport with the data source object and its name ("model").
            engine.BuildReport(loadedTemplate, new Person(), "model");
            // Save the generated report.
            loadedTemplate.Save(outputPath);

            // -----------------------------------------------------------------
            // 5. Attempt to modify restricted types AFTER a report has been built.
            //    This must throw an ArgumentException as documented.
            // -----------------------------------------------------------------
            try
            {
                // This call should fail because restricted types have already been locked.
                ReportingEngine.SetRestrictedTypes(typeof(int));
                // If no exception is thrown, indicate unexpected behavior.
                Console.WriteLine("ERROR: No exception was thrown when calling SetRestrictedTypes after BuildReport.");
            }
            catch (ArgumentException ex)
            {
                // Expected path – print confirmation.
                Console.WriteLine("Caught expected ArgumentException: " + ex.Message);
            }
            catch (Exception ex)
            {
                // Any other exception type is unexpected.
                Console.WriteLine("ERROR: Unexpected exception type: " + ex.GetType().Name);
                Console.WriteLine("Message: " + ex.Message);
            }
        }
    }
}
