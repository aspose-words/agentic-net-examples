using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample hierarchical data
        Company company = new()
        {
            Name = "Tech Corp",
            Departments = new()
            {
                new Department
                {
                    Name = "Research & Development",
                    Employees = new()
                    {
                        new Employee { Name = "Alice", Title = "Engineer" },
                        new Employee { Name = "Bob", Title = "Senior Engineer" }
                    }
                },
                new Department
                {
                    Name = "Human Resources",
                    Employees = new()
                    {
                        new Employee { Name = "Carol", Title = "HR Manager" }
                    }
                }
            }
        };

        // Create template document programmatically
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Company: <<[company.Name]>>");
        builder.Writeln("<<foreach [dept in company.Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("- <<[emp.Name]>> (<<[emp.Title]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save template to disk
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting
        Document doc = new(templatePath);

        // Enable reflection optimization
        ReportingEngine.UseReflectionOptimization = true;

        // Create reporting engine
        ReportingEngine engine = new();

        // (Optional) Register external types if the API supports it.
        // engine.RegisterExternalType("Company", typeof(Company));
        // engine.RegisterExternalType("Department", typeof(Department));
        // engine.RegisterExternalType("Employee", typeof(Employee));

        // Build the report
        engine.BuildReport(doc, company, "company");

        // Save the generated report
        const string reportPath = "Report.docx";
        doc.Save(reportPath);

        // Indicate completion
        Console.WriteLine($"Report generated: {reportPath}");
    }
}

// Data model classes
public class Company
{
    public string Name { get; set; } = "";
    public List<Department> Departments { get; set; } = new();
}

public class Department
{
    public string Name { get; set; } = "";
    public List<Employee> Employees { get; set; } = new();
}

public class Employee
{
    public string Name { get; set; } = "";
    public string Title { get; set; } = "";
}
