using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create template document
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Customer Report");
        builder.Writeln("Name: <<[root.Person.Name]>>");
        builder.Writeln("Age: <<[root.Person.Age]>>");
        builder.Writeln("Info: <<[root.Person.CustomInfo.Info]>>");

        templateDoc.Save(templatePath);

        // Load template for reporting
        Document doc = new Document(templatePath);

        // Prepare data model
        RootModel root = new()
        {
            Person = new Person
            {
                Name = "John Doe",
                Age = 30,
                CustomInfo = new CustomInfo()
            }
        };

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, root, "root");

        // Save the generated report
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {reportPath}");
    }
}

// Root wrapper class
public class RootModel
{
    public Person Person { get; set; } = new();
}

// Person data class
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public CustomInfo CustomInfo { get; set; } = new();
}

// Custom class whose members are accessed safely via a property
public class CustomInfo
{
    public string Info => "Additional custom information.";
}
