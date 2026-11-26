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
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create template document programmatically
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Person Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting
        var reportDoc = new Document(templatePath);

        // Configure ReportingEngine with known external type
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Person));

        // Build the report
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
