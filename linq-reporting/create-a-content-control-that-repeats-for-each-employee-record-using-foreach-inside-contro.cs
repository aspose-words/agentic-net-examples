using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "Alice Johnson", Age = 30, Position = "Developer" },
                new Employee { Name = "Bob Smith", Age = 45, Position = "Manager" },
                new Employee { Name = "Carol Lee", Age = 28, Position = "Designer" }
            }
        };

        // Create template document.
        const string templatePath = "EmployeeReportTemplate.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employee Report");
        builder.Writeln();

        // Begin foreach block.
        builder.Writeln("<<foreach [emp in Employees]>>");

        // Start table inside foreach.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.InsertCell();
        builder.Writeln("Position");
        builder.EndRow();

        // Data row (repeated for each employee).
        builder.InsertCell();
        builder.Writeln("<<[emp.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[emp.Age]>>");
        builder.InsertCell();
        builder.Writeln("<<[emp.Position]>>");
        builder.EndRow();

        // End table.
        builder.EndTable();

        // Close foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "EmployeeReport.docx";
        reportDoc.Save(outputPath);
    }
}
