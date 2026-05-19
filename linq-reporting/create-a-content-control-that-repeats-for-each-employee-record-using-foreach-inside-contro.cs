using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public string Position { get; set; } = "";
}

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "Alice Johnson", Position = "Developer" },
                new Employee { Name = "Bob Smith", Position = "Designer" },
                new Employee { Name = "Carol White", Position = "Manager" }
            }
        };

        // Create a template document in memory.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employees Report:");

        // Insert LINQ Reporting tags directly (no content control to avoid story boundaries).
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Position: <<[emp.Position]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("EmployeesReport.docx");
    }
}
