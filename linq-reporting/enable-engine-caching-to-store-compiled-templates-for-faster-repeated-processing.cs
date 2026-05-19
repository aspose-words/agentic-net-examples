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
        // Register code page provider required by Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Sample data.
        var persons = new List<Person>
        {
            new() { Name = "Alice", Age = 30 },
            new() { Name = "Bob", Age = 25 },
            new() { Name = "Charlie", Age = 35 }
        };

        var model = new ReportModel { Persons = persons };

        // Create a template document with LINQ Reporting tags.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        // Save and reload the template to simulate a real-world scenario.
        const string templatePath = "template.docx";
        template.Save(templatePath);
        var doc = new Document(templatePath);

        // Create a ReportingEngine instance. Reusing the same engine enables internal caching
        // of the compiled template, which speeds up subsequent BuildReport calls.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // First report generation.
        engine.BuildReport(doc, model, "model");
        doc.Save("Report1.docx");

        // Second report generation using the same engine (cached template).
        var doc2 = new Document(templatePath);
        engine.BuildReport(doc2, model, "model");
        doc2.Save("Report2.docx");
    }
}
