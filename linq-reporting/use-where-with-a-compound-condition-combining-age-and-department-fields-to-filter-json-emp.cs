using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

namespace LinqReportingJsonFilter
{
    // Data model representing an employee.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }

    // Wrapper class that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<Employee> Employees { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // 1. Prepare sample JSON data.
            string jsonPath = "employees.json";
            File.WriteAllText(jsonPath,
                @"[
                    { ""Name"": ""Alice"",   ""Age"": 28, ""Department"": ""HR"" },
                    { ""Name"": ""Bob"",     ""Age"": 35, ""Department"": ""HR"" },
                    { ""Name"": ""Charlie"", ""Age"": 42, ""Department"": ""IT"" },
                    { ""Name"": ""Diana"",   ""Age"": 31, ""Department"": ""HR"" },
                    { ""Name"": ""Evan"",    ""Age"": 25, ""Department"": ""Finance"" }
                ]");

            // 2. Load JSON into a list of Employee objects.
            List<Employee> allEmployees = JsonConvert.DeserializeObject<List<Employee>>(File.ReadAllText(jsonPath))!
                .Where(e => e != null).ToList();

            // 3. Apply a compound LINQ Where filter: Age > 30 AND Department == "HR".
            List<Employee> filteredEmployees = allEmployees
                .Where(e => e.Age > 30 && e.Department == "HR")
                .ToList();

            // 4. Prepare the model for the reporting engine.
            ReportModel model = new ReportModel { Employees = filteredEmployees };

            // 5. Create a Word document template programmatically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Employees older than 30 in HR department:");
            builder.Writeln("<<foreach [emp in Employees]>>");
            builder.Writeln("Name: <<[emp.Name]>>, Age: <<[emp.Age]>>, Dept: <<[emp.Department]>>");
            builder.Writeln("<</foreach>>");

            // 6. Build the report using Aspose.Words LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // 7. Save the generated report.
            string outputPath = "FilteredEmployeesReport.docx";
            doc.Save(outputPath);

            // Optional: clean up the temporary JSON file.
            File.Delete(jsonPath);
        }
    }
}
