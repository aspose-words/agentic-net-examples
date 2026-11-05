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
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 22 }
            }
        };

        // Create template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        // This paragraph may become empty if condition is false.
        builder.Writeln("<<if [p.Age > 30]>>Senior<</if>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template for reporting.
        var template = new Document(templatePath);

        // Configure reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(template, model, "model");

        // Save the final document.
        var outputPath = "Result.docx";
        template.Save(outputPath);
    }
}
