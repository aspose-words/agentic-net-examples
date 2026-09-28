using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ExternalInfo
{
    public bool IsActive { get; set; } = true;
}

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public ExternalInfo Info { get; set; } = new();
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();

    // Advanced filtering using Where with a lambda that references an external type property (Info.IsActive)
    public IEnumerable<Person> ActivePersons => Persons.Where(p => p.Info.IsActive);
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30, Info = new ExternalInfo { IsActive = true } },
                new Person { Name = "Bob", Age = 45, Info = new ExternalInfo { IsActive = false } },
                new Person { Name = "Charlie", Age = 25, Info = new ExternalInfo { IsActive = true } }
            }
        };

        // Create template document programmatically
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Active Persons:");
        builder.Writeln("<<foreach [p in ActivePersons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        var doc = new Document(templatePath);

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        doc.Save(outputPath);
    }
}
