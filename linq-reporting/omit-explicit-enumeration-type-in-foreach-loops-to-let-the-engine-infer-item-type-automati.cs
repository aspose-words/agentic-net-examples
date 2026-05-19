using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel();
        var persons = new List<Person>();
        foreach (var i in Enumerable.Range(1, 5))
        {
            persons.Add(new Person { Name = $"Person {i}", Age = 20 + i });
        }
        model.Persons = persons;

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        doc.Save(templatePath);

        // Load the template and build the report.
        var template = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Data model classes.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
