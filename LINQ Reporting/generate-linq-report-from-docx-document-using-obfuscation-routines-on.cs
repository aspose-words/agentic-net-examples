using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Simple POCO that will be used as a data source for the LINQ query.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template DOCX file that contains Reporting Engine tags, e.g. <<foreach [emp]>><<[Name]>> <<[Department]>> <<[Salary]:currency>><</foreach>>
            string templatePath = @"C:\Templates\EmployeeReportTemplate.docx";

            // Load the template document.
            Document templateDoc = new Document(templatePath);

            // Create a sample collection of employees.
            List<Employee> employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Department = "Finance", Salary = 72000m },
                new Employee { Name = "Jane Smith", Department = "HR", Salary = 65000m },
                new Employee { Name = "Bob Johnson", Department = "IT", Salary = 85000m },
                new Employee { Name = "Alice Brown", Department = "Marketing", Salary = 59000m }
            };

            // Use LINQ to group employees by department and calculate average salary.
            var departmentStats = employees
                .GroupBy(e => e.Department)
                .Select(g => new
                {
                    Department = g.Key,
                    EmployeeCount = g.Count(),
                    AverageSalary = g.Average(e => e.Salary),
                    Employees = g.ToList()
                })
                .ToList();

            // The ReportingEngine can work with any non‑dynamic object.
            // We expose the LINQ result as a property named "deptStats" in an anonymous wrapper.
            var dataSource = new
            {
                deptStats = departmentStats
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after tag removal.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // The template can reference the data source via the name "deptStats".
            engine.BuildReport(templateDoc, dataSource, "deptStats");

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\EmployeeReport.docx";

            // Save the populated document.
            templateDoc.Save(outputPath);

            // -----------------------------------------------------------------
            // NOTE: If you need to apply obfuscation to the target assembly
            // (e.g., the compiled EXE/DLL), you would typically run an external
            // obfuscation tool after the build step. This code does not perform
            // obfuscation itself; it only demonstrates the LINQ reporting flow.
            // -----------------------------------------------------------------
        }
    }
}
