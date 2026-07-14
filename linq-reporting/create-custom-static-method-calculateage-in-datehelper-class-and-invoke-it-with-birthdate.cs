using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        Person person = new Person
        {
            Name = "John Doe",
            BirthDate = new DateTime(1990, 5, 15)
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags.
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[DateHelper.CalculateAge(model.BirthDate)]>>");

        // Register the helper class so its static members can be used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(DateHelper));

        // Build the report using the template and the data source.
        engine.BuildReport(template, person, "model");

        // Save the generated report.
        template.Save("Report.docx");
    }
}

// Data model class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime BirthDate { get; set; }
}

// Helper class containing the custom static method.
public static class DateHelper
{
    // Calculates age based on the provided birth date.
    public static int CalculateAge(DateTime birthDate)
    {
        DateTime today = DateTime.Today;
        int age = today.Year - birthDate.Year;
        if (birthDate > today.AddYears(-age))
            age--;
        return age;
    }
}
