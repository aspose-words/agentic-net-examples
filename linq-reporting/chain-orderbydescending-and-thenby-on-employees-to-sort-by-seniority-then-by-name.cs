using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data entity representing an employee.
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
            // 1. Create sample employee data.
            var employees = new List<Employee>
            {
                new Employee { Name = "Alice",   Seniority = 5 },
                new Employee { Name = "Bob",     Seniority = 3 },
                new Employee { Name = "Charlie", Seniority = 5 },
                new Employee { Name = "David",   Seniority = 2 }
            };

            // 2. Sort by seniority (descending) then by name (ascending).
            var sortedEmployees = employees
                .OrderByDescending(e => e.Seniority)
                .ThenBy(e => e.Name)
                .ToList();

            // 3. Wrap the sorted collection in the model.
            var model = new ReportModel { Employees = sortedEmployees };

            // 4. Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            builder.Writeln("Employee Report");
            builder.Writeln();

            // LINQ Reporting foreach tag iterating over the Employees collection.
            builder.Writeln("<<foreach [emp in Employees]>>");
            // Output each employee's name and seniority.
            builder.Writeln("Name: <<[emp.Name]>>\tSeniority: <<[emp.Seniority]>>");
            builder.Writeln("<</foreach>>");

            // 5. Save the template to a local file.
            const string templatePath = "EmployeeReportTemplate.docx";
            template.Save(templatePath);

            // 6. Load the template (optional – could reuse the same Document instance).
            var doc = new Document(templatePath);

            // 7. Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model); // No root name needed; members are accessed directly.

            // 8. Save the generated report.
            const string outputPath = "EmployeeReport.docx";
            doc.Save(outputPath);
        }
    }
}
