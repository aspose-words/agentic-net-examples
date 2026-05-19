using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Person = new Person
            {
                Name = "John Doe",
                BirthDate = new DateTime(1990, 5, 15)
            }
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Name: <<[model.Person.Name]>>");
        builder.Writeln("Age: <<[DateHelper.CalculateAge(model.Person.BirthDate)]>>");
        doc.Save(templatePath);

        // Load the template and build the report.
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Register the helper class so its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(DateHelper));

        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        var outputPath = "report.docx";
        loadedDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public Person Person { get; set; } = new Person();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime BirthDate { get; set; }
}

// Helper class with the custom static method.
public static class DateHelper
{
    // Calculates age based on the provided birth date.
    public static int CalculateAge(DateTime birthDate)
    {
        var today = DateTime.Today;
        var age = today.Year - birthDate.Year;
        if (birthDate > today.AddYears(-age))
            age--;
        return age;
    }
}
