using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 }
            }
        };

        // Create a template document with a tag that serializes the model to JSON.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Serialized JSON of the model:");
        builder.Writeln("<<[model.ToJson()]>>");
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Sample data model.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();

    // Helper method to serialize the whole model to JSON.
    public string ToJson()
    {
        return JsonConvert.SerializeObject(this, Formatting.Indented);
    }
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
