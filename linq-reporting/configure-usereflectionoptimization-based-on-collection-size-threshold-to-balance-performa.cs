using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
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

public class Program
{
    // Threshold that decides whether to enable reflection optimization.
    private const int CollectionSizeThreshold = 100;

    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var model = new ReportModel();
        for (int i = 1; i <= 120; i++) // Adjust count to test both branches.
        {
            model.Persons.Add(new Person { Name = $"Person {i}", Age = 20 + i % 30 });
        }

        // Configure reflection optimization based on collection size.
        if (model.Persons.Count > CollectionSizeThreshold)
            ReportingEngine.UseReflectionOptimization = false; // Disable for large collections.
        else
            ReportingEngine.UseReflectionOptimization = true;  // Enable for small collections.

        // Create a template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        doc.Save(templatePath);

        // Load the template (demonstrates load rule usage).
        var loadedDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}
