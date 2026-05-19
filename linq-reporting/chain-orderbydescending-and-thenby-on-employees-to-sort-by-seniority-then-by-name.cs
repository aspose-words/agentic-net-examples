using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model representing an employee.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Seniority { get; set; }

        public Employee(string name, int seniority)
        {
            Name = name;
            Seniority = seniority;
        }
    }

    // Wrapper class that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Sample employee data.
            var employees = new List<Employee>
            {
                new Employee("Alice", 5),
                new Employee("Bob", 3),
                new Employee("Charlie", 5),
                new Employee("David", 2)
            };

            // Chain OrderByDescending (seniority) then ThenBy (name).
            var sortedEmployees = employees
                .OrderByDescending(e => e.Seniority)
                .ThenBy(e => e.Name)
                .ToList();

            // Prepare the model for the report.
            var model = new ReportModel { Employees = sortedEmployees };

            // Create a blank Word document.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Build the LINQ Reporting template.
            builder.Writeln("Employee Report");
            builder.Writeln("<<foreach [e in Employees]>>");
            builder.Writeln("Name: <<[e.Name]>>, Seniority: <<[e.Seniority]>>");
            builder.Writeln("<</foreach>>");

            // Generate the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the resulting document.
            doc.Save("EmployeeReport.docx");
        }
    }
}
