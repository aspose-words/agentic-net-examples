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
            // Paths for the temporary template and the generated report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple LINQ Reporting tag that references the root object.
            builder.Writeln("Report for <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Save the template to disk.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (simulating a real‑world scenario where the
            //    template is stored separately from the code).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Build the report with reflection optimization disabled.
            // -----------------------------------------------------------------
            // Preserve the original setting so we can restore it later.
            bool originalOptimization = ReportingEngine.UseReflectionOptimization;

            try
            {
                // Disable the dynamic proxy generation for this block.
                ReportingEngine.UseReflectionOptimization = false;

                // Prepare the data source.
                Person person = new Person
                {
                    Name = "Alice Smith",
                    Age = 42
                };

                // Create the reporting engine and generate the report.
                ReportingEngine engine = new ReportingEngine();
                engine.BuildReport(loadedTemplate, person, "person");
            }
            finally
            {
                // Restore the original optimization setting.
                ReportingEngine.UseReflectionOptimization = originalOptimization;
            }

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);
        }
    }
}
