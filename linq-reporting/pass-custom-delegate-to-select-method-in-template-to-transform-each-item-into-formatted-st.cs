using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;                     // Correct namespace for Encoding
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;           // Required for ShapeType if needed in future

class Program
{
    static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons =
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create template document programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report of Persons:");
        builder.Writeln("<<foreach [line in FormattedPersons]>>");
        builder.Writeln("<<[line]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Public data model classes.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();

    // Custom delegate via LINQ Select to format each person.
    public IEnumerable<string> FormattedPersons => Persons.Select(p => $"{p.Name} ({p.Age})");
}
