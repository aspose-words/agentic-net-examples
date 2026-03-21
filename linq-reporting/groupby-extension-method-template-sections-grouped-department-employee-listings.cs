using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsGroupByExample
{
    // Simple employee data class.
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }

        public Employee(string name, string department)
        {
            Name = name;
            Department = department;
        }
    }

    // Wrapper for a department group used by the reporting engine.
    public class DepartmentGroup
    {
        public string Key { get; set; }               // Department name
        public List<Employee> Items { get; set; }     // Employees in the department
    }

    // Data source class required by Aspose.Words.Reporting (must be a visible type).
    public class ReportData
    {
        public List<DepartmentGroup> Departments { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a reporting template that iterates over the pre‑grouped data.
            builder.Writeln("Employee List grouped by Department:");
            builder.Writeln("<<foreach [in Departments]>>");
            builder.Writeln("Department: <<[Key]>>");
            builder.Writeln("<<foreach [in Items]>>");
            builder.Writeln("- <<[Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("<</foreach>>");

            // Prepare sample data.
            List<Employee> employees = new List<Employee>
            {
                new Employee("John Doe", "Sales"),
                new Employee("Jane Smith", "Marketing"),
                new Employee("Bob Johnson", "Sales"),
                new Employee("Alice Brown", "HR"),
                new Employee("Tom Clark", "Marketing")
            };

            // Group employees by department and create a list of DepartmentGroup objects.
            List<DepartmentGroup> groups = employees
                .GroupBy(e => e.Department)
                .Select(g => new DepartmentGroup
                {
                    Key = g.Key,
                    Items = g.ToList()
                })
                .ToList();

            // Wrap the grouped data in a visible data source object.
            var dataSource = new ReportData { Departments = groups };

            // Build the report using the template and the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Save the populated document.
            doc.Save("EmployeeReport.docx");
        }
    }
}
