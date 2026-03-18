using System;
using System.Collections.Generic;
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

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a reporting template that uses LINQ GroupBy to group employees by department.
            // The syntax <<foreach [in Employees.GroupBy(Department)]>> creates a group for each distinct department.
            // Inside the outer loop, <<[Key]>> outputs the department name.
            // The inner loop iterates over the grouped items (employees) using the placeholder <<[Name]>>.
            builder.Writeln("Employee List grouped by Department:");
            builder.Writeln("<<foreach [in Employees.GroupBy(Department)]>>");
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

            // The reporting engine expects the data source to be an object whose members can be referenced in the template.
            // We wrap the employee list in an anonymous object with a property named "Employees".
            var dataSource = new { Employees = employees };

            // Build the report using the template and the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Save the populated document.
            doc.Save("EmployeeReport.docx");
        }
    }
}
