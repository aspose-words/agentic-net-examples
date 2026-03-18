using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingEngineExample
{
    // Sample business object
    public class Employee
    {
        public string Name { get; set; }
        public string Department { get; set; }
        public decimal Salary { get; set; }

        public Employee(string name, string department, decimal salary)
        {
            Name = name;
            Department = department;
            Salary = salary;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare a list of business objects that will serve as the primary data source.
            List<Employee> employees = new List<Employee>
            {
                new Employee("John Doe", "Sales", 56000m),
                new Employee("Jane Smith", "Marketing", 63000m),
                new Employee("Bob Johnson", "IT", 72000m)
            };

            // 2. Load the template document that contains LINQ Reporting Engine tags.
            //    Example tag in the template: <<foreach [employees]>><<[Name]>> - <<[Department]>> - <<[Salary]:currency>><</foreach>>
            Document template = new Document("Template.docx");

            // 3. Create the reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // 4. Build the report using the list as the primary data source.
            //    The object can be any non‑dynamic, non‑anonymous .NET type; a List<T> is acceptable.
            engine.BuildReport(template, employees);

            // 5. Save the generated report.
            template.Save("Report_Output.docx");
        }
    }
}
