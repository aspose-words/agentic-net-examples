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
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 42 }
            }
        };

        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tag using Where with a lambda that references an external type property.
        // The external type AgeHelper is registered later with the ReportingEngine.
        builder.Writeln("<<foreach [p in Persons.Where(p => p.Age > AgeHelper.MinAge)]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Register the external type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(AgeHelper));

        // Build the report using the model as the data source.
        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);

        // Optional: indicate completion (no interactive input).
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

// ---------------------------------------------------------------------
// External type whose static property is used inside the LINQ expression.
// ---------------------------------------------------------------------
public static class AgeHelper
{
    // This value is referenced in the template's Where clause.
    public static int MinAge { get; } = 30;
}
