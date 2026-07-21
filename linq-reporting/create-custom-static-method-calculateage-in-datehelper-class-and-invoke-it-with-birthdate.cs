using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class DateHelper
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

public class Person
{
    // Sample property used in the report.
    public DateTime BirthDate { get; set; } = DateTime.MinValue;
}

public class Program
{
    public static void Main()
    {
        // Step 1: Create a template document with a LINQ Reporting tag.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Age: <<[DateHelper.CalculateAge(BirthDate)]>>");

        // Save the template locally.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Step 2: Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Step 3: Prepare the data source.
        Person person = new Person
        {
            BirthDate = new DateTime(1990, 5, 15) // Example birth date.
        };

        // Step 4: Configure the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(DateHelper));

        // Step 5: Build the report using the data source.
        engine.BuildReport(reportDoc, person, "person");

        // Step 6: Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
