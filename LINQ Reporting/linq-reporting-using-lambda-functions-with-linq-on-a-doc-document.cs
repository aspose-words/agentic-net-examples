using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; }
    public string Department { get; set; }
    public double Salary { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and define a simple LINQ Reporting template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // The template uses the foreach tag to iterate over a collection named "employees".
        builder.Writeln("<<foreach [employees]>>");
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Dept: <<[Department]>>");
        builder.Writeln("Salary: <<[Salary]:dollarText>>");
        builder.Writeln("<<end>>");

        // Sample data source containing several employees.
        List<Employee> allEmployees = new List<Employee>
        {
            new Employee { Name = "John Doe",   Department = "HR",      Salary = 50000 },
            new Employee { Name = "Jane Smith", Department = "IT",      Salary = 75000 },
            new Employee { Name = "Bob Johnson",Department = "IT",      Salary = 65000 },
            new Employee { Name = "Alice Brown",Department = "Finance", Salary = 80000 }
        };

        // Use LINQ with a lambda expression to filter only IT department employees.
        var itEmployees = allEmployees
            .Where(e => e.Department == "IT")
            .Select(e => new { e.Name, e.Department, e.Salary })
            .ToList();

        // Build the report. The data source name "employees" must match the template tag.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, itEmployees, "employees");

        // Save the generated document.
        doc.Save("IT_Employees_Report.docx");
    }
}
