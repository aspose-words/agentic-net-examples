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

        // Create a template document using the default LINQ Reporting delimiters << and >>.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for completeness).
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting (demonstrates load‑save cycle).
        var doc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();

        // Build the report. The third parameter is the root object name used in the template.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
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
