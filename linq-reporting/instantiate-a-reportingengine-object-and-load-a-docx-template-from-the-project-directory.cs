using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Define template file path in the project directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");

        // Create a simple template if it does not exist.
        if (!File.Exists(templatePath))
        {
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            // Insert a LINQ Reporting tag that references the model.
            builder.Writeln("Hello, <<[person.Name]>>!");
            templateDoc.Save(templatePath);
        }

        // Load the template document.
        Document doc = new Document(templatePath);

        // Instantiate the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Prepare a data source.
        Person person = new Person();

        // Build the report using the loaded template and the data source.
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
