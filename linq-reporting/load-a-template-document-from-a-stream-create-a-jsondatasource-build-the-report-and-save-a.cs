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
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document in memory.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title.
        builder.Writeln("Person Report");
        builder.Writeln();

        // LINQ Reporting tags using the root name "model".
        builder.Writeln("<<foreach [person in model.Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to a memory stream.
        using var templateStream = new MemoryStream();
        templateDoc.Save(templateStream, SaveFormat.Docx);
        templateStream.Position = 0; // Reset for reading.

        // Load the template document for reporting.
        var reportDoc = new Document(templateStream);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice",   Age = 30 },
                new Person { Name = "Bob",     Age = 25 },
                new Person { Name = "Charlie", Age = 28 }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report as RTF.
        reportDoc.Save("Report.rtf", SaveFormat.Rtf);
    }
}
