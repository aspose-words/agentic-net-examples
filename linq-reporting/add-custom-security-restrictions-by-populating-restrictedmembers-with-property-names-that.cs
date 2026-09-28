using System;
using System.Collections.Generic;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        Person person = new()
        {
            Name = "John Doe",
            Age = 30,
            Salary = 75000m
        };

        // -----------------------------------------------------------------
        // Create a Word template with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document template = new();
        DocumentBuilder builder = new(template);
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Salary: <<[person.Salary]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new(templatePath);

        // -----------------------------------------------------------------
        // Configure the reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new();

        // Populate RestrictedMembers (if the property exists in the current
        // Aspose.Words version) using reflection to avoid compile‑time errors.
        PropertyInfo restrictedProp = typeof(ReportingEngine).GetProperty("RestrictedMembers");
        if (restrictedProp != null)
        {
            if (restrictedProp.GetValue(engine) is IList<string> restrictedList)
            {
                restrictedList.Add("Salary");
            }
        }

        // Build the report using the person object as the root named "person".
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        const string reportPath = "Report.docx";
        doc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {reportPath}");
    }
}

// Public data model class.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public decimal Salary { get; set; }
}
