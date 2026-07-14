using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new Employee { Id = 1, Name = "Alice Johnson", Position = "Developer" },
                new Employee { Id = 2, Name = "Bob Smith", Position = "Designer" },
                new Employee { Id = 3, Name = "Carol White", Position = "Manager" }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags directly into the document.
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Id: <<[emp.Id]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Position: <<[emp.Position]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
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
