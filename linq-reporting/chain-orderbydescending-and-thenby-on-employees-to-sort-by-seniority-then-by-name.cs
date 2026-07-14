using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data entity representing an employee.
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

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare sample data.
            var model = new ReportModel
            {
                Employees = new List<Employee>
                {
                    new Employee("Alice Johnson", 5),
                    new Employee("Bob Smith", 3),
                    new Employee("Charlie Davis", 5),
                    new Employee("Diana Evans", 2)
                }
            };

            // 2. Create a template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("Employee Report");
            builder.Writeln("Sorted by seniority (descending) then by name (ascending):");
            // LINQ Reporting foreach tag with chained OrderByDescending and ThenBy.
            builder.Writeln("<<foreach [emp in Employees.OrderByDescending(e => e.Seniority).ThenBy(e => e.Name)]>>");
            builder.Writeln("Name: <<[emp.Name]>> | Seniority: <<[emp.Seniority]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk before building the report (required by lifecycle rule).
            doc.Save(templatePath);

            // 3. Load the template and build the report.
            var templateDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(templateDoc, model, "model");

            // 4. Save the generated report.
            var outputPath = "EmployeeReport.docx";
            templateDoc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {outputPath}");
        }
    }
}
