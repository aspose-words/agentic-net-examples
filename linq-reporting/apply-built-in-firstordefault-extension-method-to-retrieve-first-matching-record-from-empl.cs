using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Model
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new Model
        {
            Employees = new List<Employee>
            {
                new Employee { Name = "Alice", Age = 28 },
                new Employee { Name = "Bob",   Age = 35 },
                new Employee { Name = "Carol", Age = 42 }
            }
        };

        // Create a blank document and insert a LINQ Reporting tag that uses FirstOrDefault.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        // Note: property name is case‑sensitive; use "Employees" as defined in the model.
        builder.Writeln("First employee older than 30: <<[Employees.FirstOrDefault(p => p.Age > 30).Name]>>");

        // Build the report using the model as the data source.
        var engine = new ReportingEngine();
        // No data source name is required because we reference members directly.
        engine.BuildReport(doc, model, null);

        // Save the generated document.
        doc.Save("Report.docx");
    }
}
