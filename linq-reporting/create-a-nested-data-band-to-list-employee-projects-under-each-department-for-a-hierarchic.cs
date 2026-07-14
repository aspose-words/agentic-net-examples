using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create the template document with LINQ Reporting tags.
        var templatePath = "ReportTemplate.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("Projects:");
        builder.Writeln("<<foreach [proj in dept.Projects]>>");
        builder.Writeln("- <<[proj.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // 2. Load the template for report generation.
        var doc = new Document(templatePath);

        // 3. Prepare sample data.
        var model = new ReportModel
        {
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research",
                    Projects = new List<Project>
                    {
                        new Project { Name = "AI Exploration" },
                        new Project { Name = "Quantum Computing" }
                    }
                },
                new Department
                {
                    Name = "Development",
                    Projects = new List<Project>
                    {
                        new Project { Name = "Mobile App" },
                        new Project { Name = "Web Platform" },
                        new Project { Name = "API Integration" }
                    }
                }
            }
        };

        // 4. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 5. Save the generated report.
        doc.Save("ReportResult.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Department> Departments { get; set; } = new();
}

// Department with a collection of projects.
public class Department
{
    public string Name { get; set; } = string.Empty;
    public List<Project> Projects { get; set; } = new();
}

// Simple project entity.
public class Project
{
    public string Name { get; set; } = string.Empty;
}
