using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 45 },
                new Person { Name = "Charlie", Age = 25 }
            }
        };

        // Build the template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Disable reflection optimization for this report.
        ReportingEngine.UseReflectionOptimization = false;

        // Use a disposable wrapper to modify engine settings inside a using block.
        using (var wrapper = new ReportingEngineWrapper())
        {
            // Example of modifying an engine option (optional).
            wrapper.Engine.Options = ReportBuildOptions.None;

            // Build the report.
            wrapper.Engine.BuildReport(doc, model, "model");
        }

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Simple wrapper to allow a using block for ReportingEngine.
public class ReportingEngineWrapper : IDisposable
{
    public ReportingEngine Engine { get; }

    public ReportingEngineWrapper()
    {
        Engine = new ReportingEngine();
    }

    public void Dispose()
    {
        // No unmanaged resources to release.
    }
}

// Data model classes.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
