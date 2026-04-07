using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new()
        {
            Employees = new()
            {
                new Employee { Id = 1, Name = "Alice Johnson", Position = "Developer" },
                new Employee { Id = 2, Name = "Bob Smith", Position = "Designer" },
                new Employee { Id = 3, Name = "Carol Davis", Position = "Manager" }
            }
        };

        // Create a temporary folder for files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document programmatically.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Employee Report");
        builder.Writeln("");
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("- Id: <<[emp.Id]>>, Name: <<[emp.Name]>>, Position: <<[emp.Position]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "EmployeeReport.docx");
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Employee
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
}
