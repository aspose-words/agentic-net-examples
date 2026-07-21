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
    public List<Person> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Person>
            {
                new() { Name = "Alice", Age = 30 },
                new() { Name = "Bob", Age = 25 },
                new() { Name = "Charlie", Age = 35 }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report:");
        // The foreach tag will let the engine infer the item type (Person) automatically.
        builder.Writeln("<<foreach [person in Items]>>");
        builder.Writeln("- <<[person.Name]>> is <<[person.Age]>> years old.");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);
        var engine = new ReportingEngine();

        // BuildReport with root object name "model" to match the template expressions.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
