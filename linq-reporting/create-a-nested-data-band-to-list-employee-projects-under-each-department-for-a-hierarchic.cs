using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write LINQ Reporting tags for hierarchical data: Departments -> Employees -> Projects.
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("  Employee: <<[emp.Name]>>");
        builder.Writeln("  <<foreach [proj in emp.Projects]>>");
        builder.Writeln("    - <<[proj.Name]>>");
        builder.Writeln("  <</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
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

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine("Report generated: " + outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Department> Departments { get; set; } = new();
}

// Department model.
public class Department
{
    public string Name { get; set; } = string.Empty;
    public List<Employee> Employees { get; set; } = new();
}

// Employee model.
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public List<Project> Projects { get; set; } = new();
}

// Project model.
public class Project
{
    public string Name { get; set; } = string.Empty;
}
