using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the template.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a Word template programmatically and save it to disk.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert simple LINQ Reporting tags that reference a root object named "model".
            builder.Writeln("Name: <<[model.Name]>>");
            builder.Writeln("Age: <<[model.Age]>>");

            // Save the template so it can be loaded later.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);
            Person model = new Person { Name = "John Doe", Age = 30 };

            // -----------------------------------------------------------------
            // 3. Disable reflection optimization for this specific report.
            // -----------------------------------------------------------------
            // Store the original setting so we can restore it afterwards.
            bool originalOptimizationSetting = ReportingEngine.UseReflectionOptimization;

            try
            {
                // Turn off the optimization to avoid the overhead of dynamic proxy generation.
                ReportingEngine.UseReflectionOptimization = false;

                // Build the report using the disabled optimization.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(loadedTemplate, model, "model");
            }
            finally
            {
                // Restore the original optimization setting for any subsequent operations.
                ReportingEngine.UseReflectionOptimization = originalOptimizationSetting;
            }

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}
