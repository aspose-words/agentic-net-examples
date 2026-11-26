using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample employee data.
        var employees = new List<Employee>
        {
            new Employee { Id = 1, Name = "John Doe", Position = "Developer" },
            new Employee { Id = 2, Name = "Jane Smith", Position = "Designer" },
            new Employee { Id = 3, Name = "Bob Johnson", Position = "Manager" }
        };

        var reportData = new ReportData { Employees = employees };

        // Create the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employee List:");

        // Begin foreach block.
        builder.Writeln("<<foreach [emp in model.Employees]>>");

        // Table that will be repeated for each employee.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("ID");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Position");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[emp.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[emp.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[emp.Position]>>");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, reportData, "model");

        // Save the generated report.
        const string outputPath = "EmployeeReport.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Saved to {Path.GetFullPath(outputPath)}");
    }
}

// Wrapper class for the root object.
public class ReportData
{
    public List<Employee> Employees { get; set; } = new();
}

// Employee data model.
public class Employee
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Position { get; set; } = string.Empty;
}
