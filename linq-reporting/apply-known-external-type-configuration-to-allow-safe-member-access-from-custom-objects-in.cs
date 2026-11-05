using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public DateTime BirthDate { get; set; }
}

public class Model
{
    public Person Person { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new Model
        {
            Person = new Person
            {
                Name = "John Doe",
                BirthDate = new DateTime(1990, 5, 23)
            }
        };

        // Create the template document.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("Name: <<[model.Person.Name]>>");
        builder.Writeln("Year of Birth: <<[model.Person.BirthDate.Year]>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the reporting engine to allow safe access to DateTime members.
        var engine = new ReportingEngine();

        // Add DateTime as a known external type (new API uses KnownTypes collection).
        engine.KnownTypes.Add(typeof(DateTime));

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
