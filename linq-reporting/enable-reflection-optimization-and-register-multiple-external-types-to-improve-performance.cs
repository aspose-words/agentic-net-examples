using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for the ReportingEngine.
        ReportingEngine.UseReflectionOptimization = true;

        // Create a simple template document programmatically.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Prepare hierarchical data.
        var model = new Company
        {
            Name = "Contoso Ltd.",
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research",
                    Employees = new List<Employee>
                    {
                        new Employee { FullName = "Alice Johnson", Salary = 95000m },
                        new Employee { FullName = "Bob Smith", Salary = 87000m }
                    }
                },
                new Department
                {
                    Name = "Development",
                    Employees = new List<Employee>
                    {
                        new Employee { FullName = "Carol White", Salary = 105000m },
                        new Employee { FullName = "David Brown", Salary = 99000m }
                    }
                }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        // Register external types that can be used inside the template.
        engine.KnownTypes.Add(typeof(Helper));
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    // Creates a template with LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Company name.
        builder.Writeln("Company: <<[model.Name]>>");
        builder.Writeln();

        // Departments loop.
        builder.Writeln("<<foreach [dept in model.Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln();

        // Employees loop inside each department.
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("- <<[emp.FullName]>> (Salary: <<[Helper.FormatSalary(emp.Salary)]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        template.Save(path);
    }
}

// Helper class with a static method that will be called from the template.
public static class Helper
{
    public static string FormatSalary(decimal salary)
    {
        return string.Format("{0:C}", salary);
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
    public string FullName { get; set; } = string.Empty;
    public decimal Salary { get; set; }
}
