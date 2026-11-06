using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}

public class DepartmentGroup
{
    public string Name { get; set; } = "";
    public List<Person> Persons { get; set; } = new();
}

public class ReportModel
{
    public List<DepartmentGroup> Departments { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        List<Person> persons = new()
        {
            new() { Name = "Alice", Age = 30, Department = "HR" },
            new() { Name = "Bob", Age = 25, Department = "IT" },
            new() { Name = "Charlie", Age = 28, Department = "HR" },
            new() { Name = "Diana", Age = 35, Department = "Finance" },
            new() { Name = "Evan", Age = 22, Department = "IT" }
        };

        // Group persons by department.
        ReportModel model = new()
        {
            Departments = persons
                .GroupBy(p => p.Department)
                .Select(g => new DepartmentGroup
                {
                    Name = g.Key,
                    Persons = g.ToList()
                })
                .ToList()
        };

        // Create the template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("<<foreach [p in dept.Persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new()
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
