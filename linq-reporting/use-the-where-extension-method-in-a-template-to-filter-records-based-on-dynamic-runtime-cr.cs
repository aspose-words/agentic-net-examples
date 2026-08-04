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
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var persons = new List<Person>
        {
            new() { Name = "Alice", Age = 28 },
            new() { Name = "Bob", Age = 35 },
            new() { Name = "Charlie", Age = 42 },
            new() { Name = "Diana", Age = 23 }
        };

        var model = new ReportModel { Persons = persons };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a foreach tag that uses the Where extension method to filter persons older than 30.
        builder.Writeln("<<foreach [p in model.Persons.Where(p => p.Age > 30)]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template (optional, demonstrates the load step).
        var loadedDoc = new Document(templatePath);

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}
