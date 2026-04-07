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

public class Model
{
    public List<Person> Persons { get; set; } = new();
    public string? Optional { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential JSON handling.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ---------- Create the template ----------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // This paragraph may become empty if Optional is null.
        builder.Writeln("<<[model.Optional]>>");

        // Loop over the collection of persons.
        builder.Writeln("<<foreach [p in model.Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required by the lifecycle rule).
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        template.Save(templatePath);

        // ---------- First data source ----------
        var model1 = new Model
        {
            Optional = null, // Will result in an empty paragraph.
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "", Age = 25 }, // Empty name creates an empty line.
                new Person { Name = "Bob", Age = 40 }
            }
        };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Load the template and build the report for the first data source.
        Document doc1 = new Document(templatePath);
        engine.BuildReport(doc1, model1, "model");

        // Save the first report.
        string output1 = Path.Combine(Environment.CurrentDirectory, "Report1.docx");
        doc1.Save(output1);

        // ---------- Second data source ----------
        var model2 = new Model
        {
            Optional = "Header text",
            Persons = new List<Person>
            {
                new Person { Name = "Charlie", Age = 22 },
                new Person { Name = "Dana", Age = 28 }
            }
        };

        // Load a fresh copy of the template for the second run.
        Document doc2 = new Document(templatePath);
        engine.BuildReport(doc2, model2, "model");

        // Save the second report.
        string output2 = Path.Combine(Environment.CurrentDirectory, "Report2.docx");
        doc2.Save(output2);
    }
}
