using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model for an employee.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Seniority { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample employee data (unsorted).
            var employees = new List<Employee>
            {
                new Employee { Name = "Alice", Seniority = 3 },
                new Employee { Name = "Bob", Seniority = 5 },
                new Employee { Name = "Charlie", Seniority = 5 },
                new Employee { Name = "David", Seniority = 2 }
            };

            // 2. Sort by seniority descending, then by name ascending using LINQ.
            var sortedEmployees = employees
                .OrderByDescending(e => e.Seniority)
                .ThenBy(e => e.Name)
                .ToList();

            // 3. Populate the wrapper model.
            var model = new ReportModel { Employees = sortedEmployees };

            // 4. Create a Word template programmatically.
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("<<foreach [employee in Employees]>>");
            builder.Writeln("Name: <<[employee.Name]>>, Seniority: <<[employee.Seniority]>>");
            builder.Writeln("<</foreach>>");

            // 5. Save the template to disk.
            const string templatePath = "EmployeeReportTemplate.docx";
            templateDoc.Save(templatePath);

            // 6. Load the template for reporting.
            var loadedTemplate = new Document(templatePath);

            // 7. Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // 8. Save the generated report.
            const string outputPath = "EmployeeReportResult.docx";
            loadedTemplate.Save(outputPath);
        }
    }
}
