using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working folder
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // Create a simple template document with LINQ Reporting tags
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");
        templateDoc.Save(templatePath);

        // Load the template
        Document doc = new Document(templatePath);

        // Sample data model
        Person model = new Person
        {
            Name = "Alice Johnson",
            Age = 28
        };

        // Configure ReportingEngine (restricted members configuration is not required for this version)
        ReportingEngine engine = new ReportingEngine();

        // Build the report
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);

        // Indicate result (no interactive input)
        Console.WriteLine(success
            ? $"Report generated successfully at: {outputPath}"
            : "Report generation failed.");
    }
}

// Data model with public properties
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
