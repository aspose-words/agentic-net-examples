using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Department Project Report");
        builder.Writeln();

        // Outer data band – iterate over departments.
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln();

        // Inner data band – iterate over projects of the current department.
        builder.Writeln("Projects:");
        builder.Writeln("<<foreach [proj in dept.Projects]>>");
        builder.Writeln("- <<[proj.Name]>> (Budget: <<[proj.Budget]>>)");
        builder.Writeln("<</foreach>>"); // End inner foreach.
        builder.Writeln(); // Blank line between departments.
        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Build the data model.
        ReportModel model = new()
        {
            Departments = new List<Department>
            {
                new()
                {
                    Name = "Research",
                    Projects = new List<Project>
                    {
                        new() { Name = "AI Platform", Budget = 150000m },
                        new() { Name = "Quantum Computing", Budget = 250000m }
                    }
                },
                new()
                {
                    Name = "Marketing",
                    Projects = new List<Project>
                    {
                        new() { Name = "Social Media Campaign", Budget = 50000m },
                        new() { Name = "Product Launch", Budget = 80000m }
                    }
                }
            }
        };

        // Generate the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("DepartmentProjectReport.docx");

        // Optional: indicate success (no console interaction required).
        if (!success)
        {
            throw new InvalidOperationException("Report generation failed.");
        }
    }
}

// Root wrapper class – must match the name used in BuildReport ("model").
public class ReportModel
{
    public List<Department> Departments { get; set; } = new();
}

// Department class.
public class Department
{
    public string Name { get; set; } = string.Empty;
    public List<Project> Projects { get; set; } = new();
}

// Project class.
public class Project
{
    public string Name { get; set; } = string.Empty;
    public decimal Budget { get; set; }
}
