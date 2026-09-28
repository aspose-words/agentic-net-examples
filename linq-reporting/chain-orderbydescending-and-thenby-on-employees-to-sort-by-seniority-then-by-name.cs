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
        // Sample employee data.
        List<Employee> employees = new()
        {
            new Employee { Name = "Alice", Seniority = 5 },
            new Employee { Name = "Bob", Seniority = 3 },
            new Employee { Name = "Charlie", Seniority = 5 },
            new Employee { Name = "David", Seniority = 2 }
        };

        // Chain OrderByDescending and ThenBy to sort by seniority descending, then by name ascending.
        List<Employee> sortedEmployees = employees
            .OrderByDescending(e => e.Seniority)
            .ThenBy(e => e.Name)
            .ToList();

        // Prepare the model for the report.
        ReportModel model = new()
        {
            Employees = sortedEmployees
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = "Template.docx";
        DocumentBuilder builder = new();
        builder.Writeln("Employee Report");
        builder.Writeln("<<foreach [e in Employees]>>");
        builder.Writeln("Name: <<[e.Name]>>, Seniority: <<[e.Seniority]>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template document.
        Document doc = new(templatePath);

        // Build the report using Aspose.Words ReportingEngine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }
}
