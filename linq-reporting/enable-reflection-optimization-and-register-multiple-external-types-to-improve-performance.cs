using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Company: <<[model.Name]>>");
        builder.Writeln("<<foreach [dept in model.Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("- <<[emp.Name]>> (<<[emp.Position]>>), Age: <<[emp.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare hierarchical data.
        Company company = new Company
        {
            Name = "Tech Solutions Ltd.",
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research & Development",
                    Employees = new List<Employee>
                    {
                        new Employee { Name = "Alice Johnson", Age = 30, Position = "Senior Engineer" },
                        new Employee { Name = "Bob Smith", Age = 28, Position = "Engineer" }
                    }
                },
                new Department
                {
                    Name = "Human Resources",
                    Employees = new List<Employee>
                    {
                        new Employee { Name = "Carol White", Age = 35, Position = "HR Manager" },
                        new Employee { Name = "David Brown", Age = 32, Position = "Recruiter" }
                    }
                }
            }
        };

        // Create and configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, company, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class Company
{
    public string Name { get; set; } = string.Empty;
    public List<Department> Departments { get; set; } = new();
}

public class Department
{
    public string Name { get; set; } = string.Empty;
    public List<Employee> Employees { get; set; } = new();
}

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Position { get; set; } = string.Empty;
}
