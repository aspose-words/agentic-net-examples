using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public string Title { get; set; } = "";
    public string Department { get; set; } = "";
}

public class DepartmentGroup
{
    public string Name { get; set; } = "";
    public List<Employee> Employees { get; set; } = new();
}

public class ReportModel
{
    public List<DepartmentGroup> Departments { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample employee data.
        List<Employee> employees = new()
        {
            new() { Name = "Alice Johnson", Title = "Developer", Department = "IT" },
            new() { Name = "Bob Smith", Title = "System Analyst", Department = "IT" },
            new() { Name = "Carol White", Title = "HR Manager", Department = "HR" },
            new() { Name = "David Brown", Title = "Recruiter", Department = "HR" },
            new() { Name = "Eve Davis", Title = "Sales Executive", Department = "Sales" }
        };

        // Group employees by department using LINQ GroupBy.
        ReportModel model = new()
        {
            Departments = employees
                .GroupBy(e => e.Department)
                .Select(g => new DepartmentGroup
                {
                    Name = g.Key,
                    Employees = g.ToList()
                })
                .ToList()
        };

        // Create the template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Employee Report");
        builder.Writeln("");

        // Outer foreach over departments.
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Name]>>");
        builder.Writeln("");

        // Inner foreach over employees within the current department.
        builder.Writeln("<<foreach [emp in dept.Employees]>>");
        builder.Writeln("- <<[emp.Name]>> (<<[emp.Title]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save("Report.docx");
    }
}
