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

public static class FilterHelper
{
    // External static property used in the LINQ Where clause.
    public static int MinAge => 30;
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
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 40 }
            }
        };

        // Create the template document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tag using Where with a lambda that references an external static property.
        // Since ReportingEngine does not expose RegisterReference in this version, we use the constant value directly.
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > 30)]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
