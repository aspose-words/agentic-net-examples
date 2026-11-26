using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DateHelper
{
    // Calculates age based on the provided birth date and the current date.
    public static int CalculateAge(DateTime birthDate)
    {
        DateTime today = DateTime.Today;
        int age = today.Year - birthDate.Year;
        if (birthDate > today.AddYears(-age))
            age--;
        return age;
    }
}

public class Person
{
    // Birth date of the person.
    public DateTime BirthDate { get; set; } = DateTime.MinValue;

    // Age calculated via the static DateHelper method.
    public int Age => DateHelper.CalculateAge(BirthDate);
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Person's Age: <<[person.Age]>>");
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        Person person = new Person
        {
            BirthDate = new DateTime(1990, 5, 15) // Sample birth date.
        };

        // Build the report using Aspose.Words LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        reportDoc.Save(reportPath);
    }
}
