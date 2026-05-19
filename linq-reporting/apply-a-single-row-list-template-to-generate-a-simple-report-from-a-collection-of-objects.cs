using System;
using System.Collections.Generic;
using System.IO;
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
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 45 },
                new Person { Name = "Charlie", Age = 28 }
            }
        };

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Simple Person Report");
        builder.Writeln();

        // LINQ Reporting foreach tag – iterates over the collection Persons.
        builder.Writeln("<<foreach [person in Persons]>>");
        // Each iteration writes a bullet line with the person's data.
        builder.Writeln("• <<[person.Name]>> (Age: <<[person.Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        var engine = new ReportingEngine();
        // Explicitly set options (none in this simple case).
        engine.Options = ReportBuildOptions.None;

        // BuildReport overload with root object name "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
    }
}
