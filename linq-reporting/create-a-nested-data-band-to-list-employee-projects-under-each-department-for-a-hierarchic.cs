using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the example files.
        string outputFolder = "output";
        System.IO.Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        string templatePath = System.IO.Path.Combine(outputFolder, "ReportTemplate.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Company Departments Report");
        builder.Writeln();

        // Begin outer foreach for departments.
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln();

        // Begin foreach for employees within a department.
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("  Employee: <<[emp.Name]>>");
        builder.Writeln();

        // Begin foreach for projects within an employee.
        builder.Writeln("<<foreach [proj in emp.Projects]>>");
        builder.Writeln("    Project: <<[proj.Name]>>");
        builder.Writeln("<</foreach>>"); // End projects foreach

        builder.Writeln("<</foreach>>"); // End employees foreach
        builder.Writeln("<</foreach>>"); // End departments foreach

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare sample hierarchical data.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research & Development",
                    Employees = new List<Employee>
                    {
                        new Employee
                        {
                            Name = "Alice Johnson",
                            Projects = new List<Project>
                            {
                                new Project { Name = "AI Platform" },
                                new Project { Name = "Quantum Computing Study" }
                            }
                        },
                        new Employee
                        {
                            Name = "Bob Smith",
                            Projects = new List<Project>
                            {
                                new Project { Name = "Robotics Prototype" }
                            }
                        }
                    }
                },
                new Department
                {
                    Name = "Marketing",
                    Employees = new List<Employee>
                    {
                        new Employee
                        {
                            Name = "Carol White",
                            Projects = new List<Project>
                            {
                                new Project { Name = "Social Media Campaign" },
                                new Project { Name = "Brand Redesign" }
                            }
                        }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        string resultPath = System.IO.Path.Combine(outputFolder, "ReportResult.docx");
        reportDoc.Save(resultPath);
    }
}

// ---------------------------------------------------------------------
// Data model definitions (public, with initialized collections).
// ---------------------------------------------------------------------
public class ReportModel
{
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
    public List<Project> Projects { get; set; } = new();
}

public class Project
{
    public string Name { get; set; } = string.Empty;
}
