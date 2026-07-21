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
        // Prepare sample data (unsorted).
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 25 },
                new Person { Name = "Carol", Age = 35 },
                new Person { Name = "Dave",  Age = 28 }
            }
        };

        // Create a template document with LINQ Reporting tags that sort the collection inline.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Build the report using the model as the root data source named "model".
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = "Report.docx";
        doc.Save(reportPath);
    }

    private static void CreateTemplate(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Persons sorted by Age (descending):");

        // Inline sorting using OrderByDescending extension method.
        builder.Writeln("<<foreach [p in model.Persons.OrderByDescending(p => p.Age)]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
