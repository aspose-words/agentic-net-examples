using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public int Seniority { get; set; }
}

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var employees = new List<Employee>
        {
            new() { Name = "Alice", Seniority = 5 },
            new() { Name = "Bob", Seniority = 3 },
            new() { Name = "Charlie", Seniority = 5 },
            new() { Name = "David", Seniority = 2 }
        };

        // Sort by seniority descending, then by name ascending.
        var sortedEmployees = employees
            .OrderByDescending(e => e.Seniority)
            .ThenBy(e => e.Name)
            .ToList();

        // Wrap the sorted collection in a model object.
        var model = new ReportModel { Employees = sortedEmployees };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employees sorted by seniority (desc) then name (asc):");
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("<<[emp.Name]>> - Seniority: <<[emp.Seniority]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}
