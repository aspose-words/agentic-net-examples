using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    // Root model passed to the reporting engine.
    public class ReportModel
    {
        // Collection of department groups.
        public List<DepartmentGroup> Departments { get; set; } = new();
    }

    // Represents a department and its employees.
    public class DepartmentGroup
    {
        public string Department { get; set; } = string.Empty;
        public List<Employee> Employees { get; set; } = new();
    }

    // Simple employee entity.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public string Department { get; set; } = string.Empty;
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample data.
            List<Employee> employees = new()
            {
                new() { Name = "Alice Johnson", Department = "HR" },
                new() { Name = "Bob Smith", Department = "IT" },
                new() { Name = "Carol White", Department = "HR" },
                new() { Name = "David Brown", Department = "Finance" },
                new() { Name = "Eve Davis", Department = "IT" }
            };

            // 2. Group employees by department using LINQ.
            ReportModel model = new()
            {
                Departments = employees
                    .GroupBy(e => e.Department)
                    .Select(g => new DepartmentGroup
                    {
                        Department = g.Key,
                        Employees = g.ToList()
                    })
                    .ToList()
            };

            // 3. Create a template document programmatically.
            Document template = new();
            DocumentBuilder builder = new(template);

            // Title.
            builder.Writeln("Employee Directory");
            builder.Writeln();

            // Begin outer foreach over departments.
            builder.Writeln("<<foreach [dept in Departments]>>");
            builder.Writeln("Department: <<[dept.Department]>>");
            builder.Writeln();

            // Begin inner foreach over employees within the current department.
            builder.Writeln("<<foreach [emp in dept.Employees]>>");
            builder.Writeln("- <<[emp.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln();

            // End outer foreach.
            builder.Writeln("<</foreach>>");

            // 4. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(template, model, "model");

            // 5. Save the generated report.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "EmployeeReport.docx");
            template.Save(outputPath);
        }
    }
}
