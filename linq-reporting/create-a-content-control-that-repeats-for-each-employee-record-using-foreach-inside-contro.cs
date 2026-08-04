using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "EmployeeTemplate.docx";
        const string outputPath = "EmployeeReport.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var builder = new DocumentBuilder();

        // Insert a paragraph that will repeat for each employee.
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Position: <<[emp.Position]>>");
        builder.Writeln("Salary: $<<[emp.Salary]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);

        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "Alice Johnson", Position = "Software Engineer", Salary = 95000m },
                new Employee { Name = "Bob Smith", Position = "Project Manager", Salary = 105000m },
                new Employee { Name = "Carol Lee", Position = "QA Analyst", Salary = 72000m }
            }
        };

        // -----------------------------------------------------------------
        // 3. Build the report.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        Console.WriteLine(success
            ? $"Report generated successfully: {outputPath}"
            : "Report generation failed.");
    }
}

// Root data model.
public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

// Employee data class.
public class Employee
{
    public string Name { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
    public decimal Salary { get; set; }
}
