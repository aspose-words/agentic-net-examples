using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DateHelper
{
    public static int CalculateAge(DateTime birthDate)
    {
        DateTime today = DateTime.Today;
        int age = today.Year - birthDate.Year;
        if (birthDate.Date > today.AddYears(-age))
            age--;
        return age;
    }
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime BirthDate { get; set; }

    public int Age => DateHelper.CalculateAge(BirthDate);
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare template
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Prepare data
        ReportModel model = new ReportModel
        {
            Persons = new List<Person>
            {
                new Person { Name = "Alice", BirthDate = new DateTime(1990, 5, 12) },
                new Person { Name = "Bob", BirthDate = new DateTime(1985, 11, 23) },
                new Person { Name = "Charlie", BirthDate = new DateTime(2000, 2, 3) }
            }
        };

        // Build report
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save output
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
