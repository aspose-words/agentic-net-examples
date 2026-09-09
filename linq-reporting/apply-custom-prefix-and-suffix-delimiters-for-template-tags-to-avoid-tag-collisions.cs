using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

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
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document using the default delimiters << and >>.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Write a foreach block with the correct LINQ Reporting syntax.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>  Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Create the reporting engine.
        var engine = new ReportingEngine();

        // Build the report. The root object is passed with the name "model"
        // so the template can reference its members directly.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save("Report.docx");
    }
}

// Root data model referenced in the template as <<[model.Persons]>>.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data entity.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
