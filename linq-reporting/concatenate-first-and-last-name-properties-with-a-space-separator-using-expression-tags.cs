using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Initialize properties to avoid nullable warnings.
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting tag that concatenates first and last name with a space.
        builder.Writeln("Full Name: <<[model.FirstName + \" \" + model.LastName]>>");

        // Prepare sample data.
        Person model = new Person
        {
            FirstName = "John",
            LastName = "Doe"
        };

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
