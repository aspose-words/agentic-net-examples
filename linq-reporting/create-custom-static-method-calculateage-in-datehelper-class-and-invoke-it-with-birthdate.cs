using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public static class DateHelper
{
    // Calculates age based on the provided birth date.
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
}

// Wrapper class required for LINQ Reporting (cannot use anonymous types).
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var persons = new List<Person>
        {
            new Person { Name = "Alice", BirthDate = new DateTime(1990, 5, 12) },
            new Person { Name = "Bob",   BirthDate = new DateTime(1985, 11, 23) }
        };

        // Create a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[DateHelper.CalculateAge(p.BirthDate)]>>");
        builder.Writeln("<</foreach>>");

        // Save and reload the template to satisfy the lifecycle rule.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);
        var doc = new Document(templatePath);

        // Prepare the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(DateHelper));

        // Build the report using a non‑anonymous root data source.
        var model = new ReportModel { Persons = persons };
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
