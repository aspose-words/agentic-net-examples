using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DateHelper
{
    // Calculates age in years based on the provided birth date.
    public static int CalculateAge(DateTime birthDate)
    {
        var today = DateTime.Today;
        int age = today.Year - birthDate.Year;
        if (birthDate > today.AddYears(-age))
            age--;
        return age;
    }
}

public class Person
{
    public string Name { get; set; } = "";
    public DateTime BirthDate { get; set; }

    // Exposes the calculated age using the static helper method.
    public int Age => DateHelper.CalculateAge(BirthDate);
}

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var person = new Person
        {
            Name = "John Doe",
            BirthDate = new DateTime(1990, 5, 15)
        };

        // Create the LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");

        // Save the template to disk.
        const string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the data model.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, person, "model");

        // Save the generated report.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
