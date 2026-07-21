using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Project
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
}

public class Department
{
    public string Name { get; set; } = "";
    public List<Project> Projects { get; set; } = new();
}

public class ReportModel
{
    public List<Department> Departments { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Departments = new List<Department>
            {
                new Department
                {
                    Name = "Research",
                    Projects = new List<Project>
                    {
                        new Project { Name = "AI Platform", Description = "Develop AI services." },
                        new Project { Name = "Quantum Study", Description = "Explore quantum algorithms." }
                    }
                },
                new Department
                {
                    Name = "Marketing",
                    Projects = new List<Project>
                    {
                        new Project { Name = "Social Campaign", Description = "Increase brand awareness." },
                        new Project { Name = "Product Launch", Description = "Launch new product line." }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Company Projects Report");
        builder.Writeln();

        // Outer foreach for departments.
        builder.Write("<<foreach [dept in Departments]>>");
        builder.Writeln();
        builder.Write("Department: <<[dept.Name]>>");
        builder.Writeln();

        // Table for projects – created inside the outer foreach so each department gets its own table.
        builder.Write("<<foreach [proj in dept.Projects]>>");
        // Start table header (only once per department).
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Project Name");
        builder.InsertCell();
        builder.Writeln("Description");
        builder.EndRow();

        // Project rows.
        builder.InsertCell();
        builder.Writeln("<<[proj.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[proj.Description]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Write("<</foreach>>"); // End inner foreach (projects)
        builder.Writeln();

        builder.Write("<</foreach>>"); // End outer foreach (departments)

        // Save the template to disk.
        const string templatePath = "ReportTemplate.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine(success ? "Report generated successfully." : "Report generation failed.");
    }
}
