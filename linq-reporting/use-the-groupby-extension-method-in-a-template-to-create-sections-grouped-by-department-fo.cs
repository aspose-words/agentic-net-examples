using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public string Department { get; set; } = "";
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
        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "Alice Johnson", Department = "HR" },
                new Employee { Name = "Bob Smith", Department = "IT" },
                new Employee { Name = "Carol White", Department = "HR" },
                new Employee { Name = "David Brown", Department = "Finance" },
                new Employee { Name = "Eve Davis", Department = "IT" }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employee Report");
        builder.Writeln();

        // Group employees by department using LINQ GroupBy inside the template.
        builder.Writeln("<<foreach [dept in Employees.GroupBy(e => e.Department)]>>");
        builder.Writeln("Department: <<[dept.Key]>>");
        builder.Writeln("<<foreach [emp in dept]>>");
        builder.Writeln("- <<[emp.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("EmployeeReport.docx");
    }
}
