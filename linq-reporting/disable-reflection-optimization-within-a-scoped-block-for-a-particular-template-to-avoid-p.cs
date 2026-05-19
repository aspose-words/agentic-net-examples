using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple paragraph that references the data source.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Disable reflection optimization for this scoped operation.
            // -----------------------------------------------------------------
            bool previousOptimizationSetting = ReportingEngine.UseReflectionOptimization;
            ReportingEngine.UseReflectionOptimization = false; // Disable proxy generation.

            try
            {
                // Create the reporting engine.
                ReportingEngine engine = new ReportingEngine();

                // Prepare the data source.
                Person person = new Person { Name = "Alice Smith", Age = 28 };

                // Build the report. The root name must match the tag prefix used in the template.
                engine.BuildReport(loadedTemplate, person, "person");

                // Save the generated report.
                loadedTemplate.Save(reportPath);
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = previousOptimizationSetting;
            }

            // The example finishes without requiring any user interaction.
        }
    }
}
