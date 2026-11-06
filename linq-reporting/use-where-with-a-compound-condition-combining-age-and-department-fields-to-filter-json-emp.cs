using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}

public class EmployeeData
{
    public List<Employee> Employees { get; set; } = new();
}

public class FilteredData
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample JSON data file.
        string jsonPath = "employees.json";
        var sampleData = new EmployeeData
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 35, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "HR" },
                new Employee { Name = "Bob Johnson", Age = 45, Department = "Sales" },
                new Employee { Name = "Alice Brown", Age = 32, Department = "IT" }
            }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON data.
        var jsonContent = File.ReadAllText(jsonPath);
        var employeeData = JsonConvert.DeserializeObject<EmployeeData>(jsonContent) ?? new EmployeeData();

        // Apply compound filter: Age > 30 AND Department == "Sales".
        var filteredEmployees = employeeData.Employees
            .Where(e => e.Age > 30 && e.Department == "Sales")
            .ToList();

        var model = new FilteredData { Employees = filteredEmployees };

        // Create LINQ Reporting template.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Employee Report");
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Age: <<[emp.Age]>>");
        builder.Writeln("Dept: <<[emp.Department]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report using the filtered data.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
