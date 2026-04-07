using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        List<Employee> employees = new()
        {
            new Employee { Name = "Alice", Age = 28 },
            new Employee { Name = "Bob",   Age = 35 },
            new Employee { Name = "Carol", Age = 42 }
        };

        // Retrieve the first employee older than 30 using LINQ FirstOrDefault.
        Employee firstMatch = employees.FirstOrDefault(e => e.Age > 30);

        // If no match is found, create a placeholder to avoid null reference.
        if (firstMatch == null)
        {
            firstMatch = new Employee { Name = "N/A", Age = 0 };
        }

        // Create a simple Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Employee Report");
        builder.Writeln("Name: <<[emp.Name]>>");
        builder.Writeln("Age:  <<[emp.Age]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the first matching employee as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, firstMatch, "emp");

        // Save the generated report.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine("Report generated: " + outputPath);
    }
}
