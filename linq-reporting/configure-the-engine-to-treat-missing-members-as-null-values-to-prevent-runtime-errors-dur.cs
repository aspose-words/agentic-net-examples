using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public Person Person { get; set; } = new Person();
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Person = new Person { Name = "Alice", Age = 30 },
            Persons = new List<Person>
            {
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create template document
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Existing member: <<[model.Person.Name]>>");
        builder.Writeln("Missing member (treated as null): <<[model.Person.NonExisting]>>");
        builder.Writeln("<<foreach [p in model.Persons]>>");
        builder.Writeln("Person: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load template for reporting
        var reportDoc = new Document(templatePath);

        // Configure reporting engine to treat missing members as null (inline error messages prevent exceptions)
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Build the report
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
