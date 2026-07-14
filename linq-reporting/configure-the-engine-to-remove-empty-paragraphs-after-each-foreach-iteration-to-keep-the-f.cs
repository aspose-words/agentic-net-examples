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
                new Person { Name = "Bob", Age = 0 },   // This entry will produce an empty paragraph.
                new Person { Name = "Charlie", Age = 25 }
            }
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("<<foreach [p in Persons]>>");
        // Paragraph that may become empty when the condition is false.
        builder.Writeln("<<if [p.Age > 0]>><<[p.Name]>> is <<[p.Age]>> years old<</if>>");
        builder.Writeln("<</foreach>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Configure the reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        doc.Save("Report.docx");
    }
}

// Wrapper class that matches the root object name used in the template.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data model used in the foreach loop.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
