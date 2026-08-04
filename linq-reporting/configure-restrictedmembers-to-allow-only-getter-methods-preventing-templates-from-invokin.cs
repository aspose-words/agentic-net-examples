using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Getter and setter are required for the template to read the values.
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class ReportModel
{
    // Initialize the property to avoid nullable warnings.
    public Person Person { get; set; } = new Person { Name = "John Doe", Age = 42 };
}

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Name: <<[model.Person.Name]>>");
        builder.Writeln("Age: <<[model.Person.Age]>>");

        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Configure the ReportingEngine.
        // No restricted types are set so that getters can be accessed.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model and the root name "model".
        engine.BuildReport(reportDoc, new ReportModel(), "model");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
