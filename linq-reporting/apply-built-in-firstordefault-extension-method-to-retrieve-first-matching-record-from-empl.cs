using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingFirstOrDefaultExample
{
    // Simple employee data model.
    public class Employee
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model that will be passed as the root object to the reporting engine.
    public class Model
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Step 1: Create a template document with a LINQ Reporting tag that uses FirstOrDefault.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("First employee over 30 years old:");
            // The expression retrieves the first employee whose Age > 30 and outputs the Name.
            builder.Writeln("<<[model.Employees.FirstOrDefault(e => e.Age > 30).Name]>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            templateDoc.Save(templatePath);

            // Step 2: Load the template (simulating a real-world scenario where the template is read from storage).
            var doc = new Document(templatePath);

            // Step 3: Prepare sample data.
            var model = new Model
            {
                Employees = new List<Employee>
                {
                    new Employee { Name = "Alice", Age = 28 },
                    new Employee { Name = "Bob", Age = 35 },
                    new Employee { Name = "Charlie", Age = 42 }
                }
            };

            // Step 4: Build the report using Aspose.Words ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Step 5: Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
