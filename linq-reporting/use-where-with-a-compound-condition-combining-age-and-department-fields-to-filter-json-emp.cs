using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data
        string jsonPath = "employees.json";
        string jsonContent = @"
[
    { ""Name"": ""Alice Johnson"", ""Age"": 28, ""Department"": ""HR"" },
    { ""Name"": ""Bob Smith"", ""Age"": 35, ""Department"": ""Sales"" },
    { ""Name"": ""Carol White"", ""Age"": 42, ""Department"": ""Sales"" },
    { ""Name"": ""David Brown"", ""Age"": 31, ""Department"": ""IT"" },
    { ""Name"": ""Eve Davis"", ""Age"": 45, ""Department"": ""Sales"" }
]";
        File.WriteAllText(jsonPath, jsonContent.Trim());

        // Deserialize JSON to list of employees
        List<Employee> allEmployees = JsonConvert.DeserializeObject<List<Employee>>(File.ReadAllText(jsonPath)) ?? new List<Employee>();

        // Filter employees: Age > 30 AND Department == "Sales"
        List<Employee> filteredEmployees = allEmployees
            .Where(e => e.Age > 30 && e.Department == "Sales")
            .ToList();

        // Prepare the model for the report
        ReportModel model = new()
        {
            Employees = filteredEmployees
        };

        // Create a Word template with LINQ Reporting tags
        string templatePath = "EmployeeReportTemplate.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Employee Report - Filtered List");
        builder.Writeln("-------------------------------------------------");
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Age: <<[emp.Age]>>");
        builder.Writeln("Department: <<[emp.Department]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("-------------------------------------------------");

        templateDoc.Save(templatePath);

        // Load the template document
        Document doc = new(templatePath);

        // Build the report
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "EmployeeReport.docx";
        doc.Save(outputPath);

        // Clean up temporary files (optional)
        // File.Delete(jsonPath);
        // File.Delete(templatePath);
    }
}

// Data model classes
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Department { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}
