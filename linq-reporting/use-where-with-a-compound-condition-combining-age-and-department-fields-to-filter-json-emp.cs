using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace AsposeWordsLinqReportingExample
{
    // Model representing an employee record.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }

    // Wrapper class exposing the collection to the template.
    public class EmployeesRoot
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for any legacy encodings.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // -----------------------------------------------------------------
            // 1. Create sample JSON data file with employee records.
            // -----------------------------------------------------------------
            string jsonPath = "employees.json";
            var sampleEmployees = new List<Employee>
            {
                new Employee { Name = "Alice", Age = 28, Department = "HR" },
                new Employee { Name = "Bob",   Age = 35, Department = "IT" },
                new Employee { Name = "Carol", Age = 42, Department = "HR" },
                new Employee { Name = "Dave",  Age = 31, Department = "Finance" },
                new Employee { Name = "Eve",   Age = 45, Department = "HR" }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleEmployees, Formatting.Indented));

            // -----------------------------------------------------------------
            // 2. Load JSON, filter with a compound LINQ condition (Age > 30 && Department == "HR").
            // -----------------------------------------------------------------
            var allEmployees = JsonConvert.DeserializeObject<List<Employee>>(File.ReadAllText(jsonPath)) ?? new List<Employee>();
            var filteredEmployees = allEmployees
                .Where(e => e.Age > 30 && e.Department == "HR")
                .ToList();

            var dataRoot = new EmployeesRoot { Employees = filteredEmployees };

            // -----------------------------------------------------------------
            // 3. Build the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            string templatePath = "EmployeeReportTemplate.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Employee Report");
            builder.Writeln("<<foreach [emp in Employees]>>");
            builder.Writeln("Name: <<[emp.Name]>>, Age: <<[emp.Age]>>, Dept: <<[emp.Department]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 4. Load the template and generate the report using the filtered data.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, dataRoot); // No root name needed; members are accessed directly.

            // -----------------------------------------------------------------
            // 5. Save the final report.
            // -----------------------------------------------------------------
            string outputPath = "EmployeeReport.docx";
            reportDoc.Save(outputPath);

            // Indicate completion (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }
    }
}
