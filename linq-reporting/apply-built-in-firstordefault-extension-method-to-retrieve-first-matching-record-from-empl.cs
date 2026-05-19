using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
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
        var model = new ReportModel();
        model.Employees.Add(new Employee { Name = "Alice", Age = 28, Department = "HR" });
        model.Employees.Add(new Employee { Name = "Bob", Age = 35, Department = "IT" });
        model.Employees.Add(new Employee { Name = "Charlie", Age = 42, Department = "Finance" });

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("First employee older than 30: <<[model.Employees.FirstOrDefault(p => p.Age > 30).Name]>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("FirstOrDefaultReport.docx");
    }
}
