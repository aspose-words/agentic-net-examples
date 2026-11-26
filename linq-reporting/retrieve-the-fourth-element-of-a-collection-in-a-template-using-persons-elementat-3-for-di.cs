using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 28 },
                new Person { Name = "Diana", Age = 32 },
                new Person { Name = "Ethan", Age = 27 }
            }
        };

        // Create template document with LINQ Reporting tag to fetch the fourth element (index 3)
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Fourth person name: <<[model.Persons.ElementAt(3).Name]>>");
        doc.Save(templatePath);

        // Load the template for reporting
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        var outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
