using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Public properties must be initialized to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Create a simple template document with LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");

        // Disable reflection optimization for small data sets.
        ReportingEngine.UseReflectionOptimization = false;

        // Prepare sample data.
        Person person = new Person { Name = "John Doe", Age = 30 };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
