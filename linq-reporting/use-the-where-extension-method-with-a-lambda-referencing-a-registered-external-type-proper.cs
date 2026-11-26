using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory("output");

        // 1. Create sample data.
        var model = new ReportModel
        {
            People = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob", Age = 35 },
                new Person { Name = "Charlie", Age = 42 },
                new Person { Name = "Diana", Age = 31 }
            }
        };

        // 2. Create a Word template programmatically.
        var templatePath = Path.Combine("output", "template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("People older than 30 (filtered using ExternalHelper.IsActive):");
        builder.Writeln("<<foreach [p in People]>>");
        builder.Writeln("- <<[p.Name]>> (Age: <<[p.Age]>>)");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // 3. Load the template.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // 4. Apply advanced filtering using Where with a lambda that references the external type.
        // The lambda checks ExternalHelper.IsActive and filters by Age > 30.
        if (ExternalHelper.IsActive)
        {
            model.People = model.People
                .Where(p => p.Age > 30)
                .ToList();
        }

        // 5. Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated report.
        var outputPath = Path.Combine("output", "report.docx");
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generated at: {outputPath}");
    }
}

// External helper class whose static property is used in filtering.
public static class ExternalHelper
{
    // This could be toggled to enable/disable filtering globally.
    public static bool IsActive { get; set; } = true;
}

// Root data model.
public class ReportModel
{
    public List<Person> People { get; set; } = new();
}

// Simple data class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
