using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple POCO class that will be used as a data source for the report.
    public class Employee
    {
        public string Name { get; set; }
        public string Position { get; set; }
        public decimal Salary { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCM template that contains the reporting tags.
            // The template must be a macro‑enabled document (DOCM) because
            // the ReportingEngine can work with any Word format.
            string templatePath = @"C:\Docs\ReportTemplate.docm";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a collection of employees that will be bound to the template.
            // The ReportingEngine can work with any IEnumerable<T> source, including LINQ queries.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Position = "Developer", Salary = 75000m },
                new Employee { Name = "Jane Smith", Position = "Designer", Salary = 68000m },
                new Employee { Name = "Bob Johnson", Position = "Manager", Salary = 92000m }
            };

            // Example LINQ query – order employees by salary descending.
            var orderedEmployees = employees.OrderByDescending(e => e.Salary);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The data source name ("employees") must match the name used
            // inside the template tags, e.g. <<foreach [employees]>><<[Name]>>...
            engine.BuildReport(doc, orderedEmployees, "employees");

            // Save the populated document. The output format is inferred from the extension.
            string outputPath = @"C:\Docs\GeneratedReport.docx";
            doc.Save(outputPath);
        }
    }
}
