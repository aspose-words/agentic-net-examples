using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare template document.
        var templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("  Employee: <<[emp.Name]>>");
        builder.Writeln("  <<foreach [proj in emp.Projects]>>");
        builder.Writeln("    Project: <<[proj.Name]>>");
        builder.Writeln("  <</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load template for reporting.
        var doc = new Document(templatePath);

        // Create sample data.
        var model = new ReportModel
        {
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research",
                    Employees = new List<Employee>
                    {
                        new Employee
                        {
                            Name = "Alice",
                            Projects = new List<Project>
                            {
                                new Project { Name = "AI Platform" },
                                new Project { Name = "Data Mining" }
                            }
                        },
                        new Employee
                        {
                            Name = "Bob",
                            Projects = new List<Project>
                            {
                                new Project { Name = "Quantum Computing" }
                            }
                        }
                    }
                },
                new Department
                {
                    Name = "Development",
                    Employees = new List<Employee>
                    {
                        new Employee
                        {
                            Name = "Charlie",
                            Projects = new List<Project>
                            {
                                new Project { Name = "Mobile App" },
                                new Project { Name = "Web Portal" }
                            }
                        }
                    }
                }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Root model class.
public class ReportModel
{
    public List<Department> Departments { get; set; } = new();
}

// Department class.
public class Department
{
    public string Name { get; set; } = "";
    public List<Employee> Employees { get; set; } = new();
}

// Employee class.
public class Employee
{
    public string Name { get; set; } = "";
    public List<Project> Projects { get; set; } = new();
}

// Project class.
public class Project
{
    public string Name { get; set; } = "";
}
