using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Employee
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Position { get; set; } = "";
}

public class ReportModel
{
    public List<Employee> Employees { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Employees = new List<Employee>
            {
                new() { Name = "Alice", Age = 28, Position = "Developer" },
                new() { Name = "Bob", Age = 35, Position = "Manager" },
                new() { Name = "Charlie", Age = 42, Position = "Director" }
            }
        };

        // Create a template document with a LINQ Reporting tag that uses FirstOrDefault.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("First employee over 30: <<[Employees.FirstOrDefault(p => p.Age > 30).Name]>>");
        builder.Document.Save(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
