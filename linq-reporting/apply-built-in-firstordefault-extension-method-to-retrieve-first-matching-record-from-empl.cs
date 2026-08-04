using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Department { get; set; } = string.Empty;
}

public class ReportModel
{
    public Employee Employee { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data source.
        List<Employee> employees = new()
        {
            new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
            new Employee { Name = "Jane Smith", Age = 25, Department = "HR" },
            new Employee { Name = "Bob Johnson", Age = 40, Department = "IT" }
        };

        // Retrieve the first employee whose name contains "John".
        Employee? firstMatch = employees.FirstOrDefault(e => e.Name.Contains("John"));

        if (firstMatch != null)
        {
            Console.WriteLine($"First matching employee: {firstMatch.Name}, Age {firstMatch.Age}, Dept {firstMatch.Department}");
        }
        else
        {
            Console.WriteLine("No matching employee found.");
        }

        // Create a simple Word template programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use the full path to the model property in the template tags.
        builder.Writeln("First matching employee: <<[model.Employee.Name]>> (Age: <<[model.Employee.Age]>>), Dept: <<[model.Employee.Department]>>");

        // Prepare the model for the reporting engine.
        ReportModel model = new()
        {
            Employee = firstMatch ?? new Employee()
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("FirstEmployeeReport.docx");
    }
}
