using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model with a birth date.
        Person person = new Person
        {
            Name = "John Doe",
            BirthDate = new DateTime(1990, 5, 15)
        };

        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a simple sentence that uses the custom static method to calculate age.
        builder.Writeln("Name: <<[Name]>>");
        builder.Writeln("Age: <<[DateHelper.CalculateAge(BirthDate)]>>");

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // Register the helper class so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(DateHelper));

        // Build the report using the person object as the root data source named "model".
        engine.BuildReport(doc, person, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Simple data model class.
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
        // Adjust if the birthday hasn't occurred yet this year.
        if (birthDate > today.AddYears(-age))
            age--;
        return age;
    }
}
