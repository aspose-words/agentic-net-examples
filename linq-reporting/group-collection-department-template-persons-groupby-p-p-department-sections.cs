using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsGroupingExample
{
    // Simple data class representing a person with a department.
    public class Employee
    {
        public string FullName { get; set; }
        public string Department { get; set; }

        public Employee(string fullName, string department)
        {
            FullName = fullName;
            Department = department;
        }
    }

    // Helper class that holds a department name and the list of employees belonging to it.
    public class DepartmentGroup
    {
        public string Department { get; set; }
        public List<Employee> Persons { get; set; }

        public DepartmentGroup(string department, List<Employee> persons)
        {
            Department = department;
            Persons = persons;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare a flat collection of employees.
            List<Employee> persons = new List<Employee>
            {
                new Employee("John Doe", "Sales"),
                new Employee("Jane Smith", "Marketing"),
                new Employee("Bob Johnson", "Sales"),
                new Employee("Alice Brown", "HR"),
                new Employee("Tom Clark", "Marketing")
            };

            // 2. Group the collection by Department using LINQ.
            var grouped = persons
                .GroupBy(p => p.Department)
                .Select(g => new DepartmentGroup(g.Key, g.ToList()))
                .ToList();

            // 3. Create a template document that contains ReportingEngine tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Begin outer foreach over department groups.
            builder.Writeln("<<foreach [dept in Departments]>>");
            // Output department name.
            builder.Writeln("Department: <<[dept.Department]>>");
            // Begin inner foreach over persons within the department.
            builder.Writeln("<<foreach [person in dept.Persons]>>");
            // Output person name.
            builder.Writeln("- <<[person.FullName]>>");
            // End inner foreach.
            builder.Writeln("<<</foreach>>");
            // Add a blank line between departments.
            builder.Writeln("");
            // End outer foreach.
            builder.Writeln("<<</foreach>>");

            // 4. Build the report by supplying the grouped data as the data source.
            ReportingEngine engine = new ReportingEngine();
            // The anonymous object exposes a property named "Departments" that matches the tag used in the template.
            engine.BuildReport(template, new { Departments = grouped });

            // 5. Save the resulting document.
            template.Save("GroupedReport.docx");
        }
    }
}
