using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Department { get; set; } = "";
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        var persons = new List<Person>
        {
            new Person { Name = "Alice", Age = 30, Department = "HR" },
            new Person { Name = "Bob", Age = 45, Department = "Finance" },
            new Person { Name = "Charlie", Age = 28, Department = "HR" },
            new Person { Name = "Diana", Age = 35, Department = "IT" },
            new Person { Name = "Ethan", Age = 40, Department = "Finance" }
        };

        var model = new ReportModel { Persons = persons };

        // Create template document.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employees Report");
        builder.Writeln("");

        // Grouping by Department.
        builder.Writeln("<<foreach [deptGroup in model.Persons.GroupBy(p => p.Department)]>>");
        builder.Writeln("Department: <<[deptGroup.Key]>>");
        builder.Writeln("<<foreach [p in deptGroup]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template and build report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        bool success = engine.BuildReport(reportDoc, model, "model");

        if (success)
        {
            var outputPath = "report.docx";
            reportDoc.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
        else
        {
            Console.WriteLine("Report generation failed.");
        }
    }
}
