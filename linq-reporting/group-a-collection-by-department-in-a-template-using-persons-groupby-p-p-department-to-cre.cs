using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        // Sample data
        var persons = new List<Person>
        {
            new Person { Name = "Alice", Age = 30, Department = "HR" },
            new Person { Name = "Bob", Age = 25, Department = "IT" },
            new Person { Name = "Charlie", Age = 28, Department = "HR" },
            new Person { Name = "Diana", Age = 35, Department = "Finance" },
            new Person { Name = "Evan", Age = 40, Department = "IT" }
        };

        var model = new ReportModel { Persons = persons };

        // Create template document
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Employee Report");
        builder.Writeln();

        // Group by Department
        builder.Writeln("<<foreach [dept in model.Persons.GroupBy(p => p.Department)]>>");
        builder.Writeln("Department: <<[dept.Key]>>");
        builder.Writeln();

        // List persons in each department
        builder.Writeln("<<foreach [person in dept]>>");
        builder.Writeln("- <<[person.Name]>> (Age: <<[person.Age]>>)");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template and build report
        var template = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save final report
        var outputPath = "Report.docx";
        template.Save(outputPath);
    }
}
