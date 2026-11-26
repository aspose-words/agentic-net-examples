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
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
    public int MinAge { get; set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Sample data
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 40 }
            }
        };

        // Create template
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln($"Persons aged >= {model.MinAge}:");
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age >= MinAge)]>>");
        builder.Writeln("- <<[p.Name]>> (<<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Build report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        engine.BuildReport(reportDoc, model, "model");

        // Save output
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
