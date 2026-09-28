using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}

public class DepartmentGroup
{
    public string Department { get; set; } = "";
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
        // Sample data.
        List<Person> persons = new()
        {
            new Person { Name = "Alice", Age = 30, Department = "HR" },
            new Person { Name = "Bob", Age = 25, Department = "IT" },
            new Person { Name = "Charlie", Age = 28, Department = "HR" },
            new Person { Name = "Diana", Age = 35, Department = "Finance" },
            new Person { Name = "Evan", Age = 22, Department = "IT" }
        };

        // Group by department.
        ReportModel model = new()
        {
            Departments = persons
                .GroupBy(p => p.Department)
                .Select(g => new DepartmentGroup
                {
                    Department = g.Key,
                    Persons = g.ToList()
                })
                .ToList()
        };

        // Create template.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Employees grouped by Department");
        builder.Writeln();

        // Outer foreach over departments.
        builder.Writeln("<<foreach [dept in Departments]>>");
        builder.Writeln("Department: <<[dept.Department]>>");
        builder.Writeln();

        // Inner foreach over persons – each person gets its own table (avoids table state issues).
        builder.Writeln("<<foreach [p in dept.Persons]>>");

        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Age]>>");
        builder.EndRow();

        builder.EndTable();

        builder.Writeln("<</foreach>>"); // End inner foreach.
        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save template.
        templateDoc.Save(templatePath);

        // Load template and build report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
