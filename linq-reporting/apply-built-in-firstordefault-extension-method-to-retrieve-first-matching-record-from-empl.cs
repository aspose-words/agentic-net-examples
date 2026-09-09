using System;
using System.Collections.Generic;
using System.Linq;
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
        // Prepare sample data.
        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "Alice", Age = 28 },
                new Employee { Name = "Bob",   Age = 35 },
                new Employee { Name = "Carol", Age = 42 }
            }
        };

        // Create a template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First employee over 30: <<[model.Employees.FirstOrDefault(p => p.Age > 30).Name]>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("FirstOrDefaultReport.docx");
    }
}
