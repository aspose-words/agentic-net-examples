using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a simple foreach loop that accesses a possibly missing property.
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>"); // Age may be null.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = null }, // Missing Age.
                new Person { Name = "Charlie", Age = 25 }
            }
        };

        // Configure the reporting engine.
        // By default, missing members are treated as null, so no special option is required.
        ReportingEngine engine = new();

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Sample item class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int? Age { get; set; }
}
