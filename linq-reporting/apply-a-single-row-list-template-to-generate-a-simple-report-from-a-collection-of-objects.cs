using System;
using System.Collections.Generic;
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
        // Sample data
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new() { Name = "Alice", Age = 30 },
                new() { Name = "Bob", Age = 25 },
                new() { Name = "Charlie", Age = 35 }
            }
        };

        // Create template document with LINQ Reporting tags
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Person Report");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template
        var doc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        const string reportPath = "report.docx";
        doc.Save(reportPath);
    }
}
