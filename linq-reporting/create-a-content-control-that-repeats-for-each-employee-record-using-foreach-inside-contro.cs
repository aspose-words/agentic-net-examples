using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var builder = new DocumentBuilder();

        // Title paragraph.
        builder.Writeln("Employee List:");

        // LINQ Reporting tags placed directly after the title.
        builder.Writeln("<<foreach [emp in Employees]>>");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Age: <<[emp.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        builder.Document.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Employees = new()
            {
                new Employee { Name = "Alice Johnson", Age = 30 },
                new Employee { Name = "Bob Smith", Age = 45 },
                new Employee { Name = "Carol Davis", Age = 28 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
