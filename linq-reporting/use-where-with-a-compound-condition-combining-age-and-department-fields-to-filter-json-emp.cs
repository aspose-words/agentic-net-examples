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

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON employee data.
        string json = @"
[
  { ""Name"": ""Alice"", ""Age"": 28, ""Department"": ""HR"" },
  { ""Name"": ""Bob"", ""Age"": 35, ""Department"": ""Sales"" },
  { ""Name"": ""Charlie"", ""Age"": 42, ""Department"": ""Sales"" },
  { ""Name"": ""Diana"", ""Age"": 31, ""Department"": ""IT"" }
]";

        // Deserialize JSON into a list of Employee objects.
        List<Employee> allEmployees = JsonConvert.DeserializeObject<List<Employee>>(json) ?? new();

        // Filter employees where Age > 30 and Department == "Sales".
        List<Employee> filtered = allEmployees
            .Where(e => e.Age > 30 && e.Department == "Sales")
            .ToList();

        // Prepare the model for the report.
        ReportModel model = new() { Employees = filtered };

        // Create the LINQ Reporting template.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Employees (Age > 30 && Department = Sales):");
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>, Age: <<[emp.Age]>>, Dept: <<[emp.Department]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template and generate the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
