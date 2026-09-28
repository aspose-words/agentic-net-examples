using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Nullable property that may be null in the data.
    public string? Name { get; set; }

    // Non‑nullable property.
    public int Age { get; set; }

    // Initialize to avoid nullable warnings for non‑nullable members.
    public Person()
    {
        Name = string.Empty;
        Age = 0;
    }
}

public class ReportModel
{
    // Collection of persons to iterate over in the template.
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        builder.Writeln("<<foreach [p in Persons]>>");
        // Use the null‑coalescing operator to provide a fallback when Name is null.
        builder.Writeln("Name: <<[p.Name ?? \"(no name)\"]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // Step 2: Load the template for report generation.
        var doc = new Document(templatePath);

        // Step 3: Prepare sample data with a null Name value.
        var model = new ReportModel
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = null, Age = 25 }, // This entry will trigger the fallback text.
                new Person { Name = "Bob", Age = 40 }
            }
        };

        // Step 4: Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Step 5: Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
