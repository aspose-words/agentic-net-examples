using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

// Register code page provider for potential encoding needs.
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

namespace AsposeWordsLinqReportingExample
{
    // Data model representing an employee.
    public class Employee
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string Department { get; set; } = string.Empty;
    }

    // Wrapper class used as the root data source for the report.
    public class ReportData
    {
        public List<Employee> Employees { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the sample JSON, template, and output documents.
            string jsonPath = "employees.json";
            string templatePath = "template.docx";
            string outputPath = "Report.docx";

            // 1. Create sample JSON data.
            CreateSampleJson(jsonPath);

            // 2. Load JSON and filter records using a compound LINQ condition.
            List<Employee> allEmployees = JsonConvert.DeserializeObject<List<Employee>>(File.ReadAllText(jsonPath))!;
            List<Employee> filteredEmployees = allEmployees
                .Where(e => e.Age > 30 && e.Department == "HR")
                .ToList();

            // 3. Prepare the root data object for the reporting engine.
            ReportData data = new ReportData { Employees = filteredEmployees };

            // 4. Build the template document with LINQ Reporting tags.
            CreateTemplateDocument(templatePath);

            // 5. Load the template, run the report, and save the result.
            Document doc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;
            engine.BuildReport(doc, data, "model");
            doc.Save(outputPath);
        }

        // Generates a simple JSON file containing a list of employees.
        private static void CreateSampleJson(string path)
        {
            var sample = new List<Employee>
            {
                new Employee { Name = "Alice Johnson", Age = 28, Department = "HR" },
                new Employee { Name = "Bob Smith", Age = 45, Department = "HR" },
                new Employee { Name = "Carol White", Age = 35, Department = "Finance" },
                new Employee { Name = "David Brown", Age = 50, Department = "HR" },
                new Employee { Name = "Eve Davis", Age = 32, Department = "IT" }
            };

            string json = JsonConvert.SerializeObject(sample, Formatting.Indented);
            File.WriteAllText(path, json);
        }

        // Creates a Word template containing LINQ Reporting tags.
        private static void CreateTemplateDocument(string path)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Employees over 30 in HR department:");
            builder.Writeln("<<foreach [emp in model.Employees]>>");
            builder.Writeln("- <<[emp.Name]>> (Age: <<[emp.Age]>>, Dept: <<[emp.Department]>>)");
            builder.Writeln("<</foreach>>");

            doc.Save(path);
        }
    }
}
