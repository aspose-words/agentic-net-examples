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
        model.Employees.Add(new Employee { Name = "Alice", Age = 28 });
        model.Employees.Add(new Employee { Name = "Bob", Age = 35 });
        model.Employees.Add(new Employee { Name = "Charlie", Age = 42 });

        // Create a template document with a LINQ Reporting tag that uses FirstOrDefault.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("First employee over 30: <<[model.Employees.FirstOrDefault(p => p.Age > 30).Name]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
