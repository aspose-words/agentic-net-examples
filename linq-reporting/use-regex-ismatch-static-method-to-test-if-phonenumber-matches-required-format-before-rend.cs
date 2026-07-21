using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        var person = new Person
        {
            Name = "John Doe",
            PhoneNumber = "123-456-7890"
        };

        // Build the template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Write static text and LINQ Reporting tags.
        builder.Writeln("Name: <<[person.Name]>>");

        // Conditional block that checks the phone number format using Regex.IsMatch.
        builder.Writeln("<<if [Regex.IsMatch(person.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>><<[person.PhoneNumber]>> (valid)<</if>>");
        builder.Writeln("<<if [!Regex.IsMatch(person.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>><<[person.PhoneNumber]>> (invalid)<</if>>");

        // Save the template to a temporary file.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Regex));
        engine.Options = ReportBuildOptions.None;

        // Build the report using the data model.
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model class with non‑nullable properties.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public string PhoneNumber { get; set; } = string.Empty;
}
