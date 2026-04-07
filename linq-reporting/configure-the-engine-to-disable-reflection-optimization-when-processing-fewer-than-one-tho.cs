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
        // 1. Prepare sample data (fewer than 1000 records).
        var model = new ReportModel();
        model.Persons.AddRange(new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        });

        // 2. Create a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // 3. Save and reload the template (required lifecycle step).
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);
        var loadedTemplate = new Document(templatePath);

        // 4. Disable reflection optimization for small data sets.
        ReportingEngine.UseReflectionOptimization = false;

        // 5. Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // 6. Save the generated report.
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
