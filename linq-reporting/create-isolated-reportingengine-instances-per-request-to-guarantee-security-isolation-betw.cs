using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(baseDir, "Template.docx");
        string outputDir = Path.Combine(baseDir, "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add a simple heading.
        builder.Writeln("User Report");
        builder.Writeln();

        // Insert a foreach tag that iterates over a collection named Persons.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>   Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // -----------------------------------------------------------------
        // 3. Simulate multiple isolated reporting requests.
        // -----------------------------------------------------------------
        for (int requestId = 1; requestId <= 3; requestId++)
        {
            // Each request gets its own ReportingEngine instance.
            ReportingEngine engine = new();

            // Load a fresh copy of the template for this request.
            Document reportDoc = new Document(templatePath);

            // Build the report using the isolated engine.
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            string reportPath = Path.Combine(outputDir, $"Report_{requestId}.docx");
            reportDoc.Save(reportPath);
        }
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
