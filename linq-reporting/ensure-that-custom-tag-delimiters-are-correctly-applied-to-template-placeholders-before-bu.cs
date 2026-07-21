using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document using the default LINQ Reporting delimiters << and >>.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>  Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        template.Save(outputPath);
    }
}

// Wrapper class that matches the root object name used in BuildReport.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple data model class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
