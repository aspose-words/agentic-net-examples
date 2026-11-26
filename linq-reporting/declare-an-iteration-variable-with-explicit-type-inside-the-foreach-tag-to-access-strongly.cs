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
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data
        ReportModel model = new()
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            }
        };

        // Create a template document programmatically
        string templatePath = "Template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Persons Report");
        // Correct foreach syntax with explicit iteration variable type omitted
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template (optional, can use the same document instance)
        Document doc = new Document(templatePath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = "Report.docx";
        doc.Save(outputPath);

        // Indicate result (no interactive input)
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
