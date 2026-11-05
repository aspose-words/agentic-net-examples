using System;
using System.Collections.Generic;
using System.IO;
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
                new Person { Name = "Alice Johnson", Age = 28 },
                new Person { Name = "Bob Smith", Age = 35 },
                new Person { Name = "Carol Davis", Age = 42 }
            }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("People Report");
        builder.Writeln("==============");
        builder.Writeln();
        // Data band: foreach over Persons collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        // Inside the band we refer to members without object reference.
        builder.Writeln("Name: <<[Name]>>, Age: <<[Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        Directory.CreateDirectory("Output");
        const string outputPath = "Output/Report.docx";
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Item class used in the data band.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
