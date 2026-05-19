using System;
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
            // Create a new blank document and a builder to insert LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple template that references the root object "person".
            builder.Writeln("<<[person.Name]>> is <<[person.Age]>> years old.");

            // Save the template to a local file (optional, demonstrates load/save workflow).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // Load the template back (simulating a separate load step).
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            Person person = new Person { Name = "Alice Smith", Age = 28 };

            // Store the original reflection optimization setting.
            bool originalOptimization = ReportingEngine.UseReflectionOptimization;

            try
            {
                // Disable reflection optimization for this report generation.
                ReportingEngine.UseReflectionOptimization = false;

                // Build the report using the LINQ Reporting engine.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(loadedTemplate, person, "person");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = originalOptimization;
            }

            // Save the generated report.
            const string outputPath = "Report.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
