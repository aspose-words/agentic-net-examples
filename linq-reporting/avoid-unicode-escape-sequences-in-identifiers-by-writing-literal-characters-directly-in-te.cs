using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Имя { get; set; } = "";
    public int Возраст { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data with Unicode property names.
        var person = new Person
        {
            Имя = "Иван Иванов",
            Возраст = 30
        };

        // Define file paths.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // Create the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Имя клиента: <<[person.Имя]>>");
        builder.Writeln("Возраст: <<[person.Возраст]>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Build the report using LINQ Reporting Engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report.
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
