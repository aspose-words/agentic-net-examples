using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class PhoneNumberModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string PhoneNumber { get; set; } = string.Empty;

    // Determines whether the phone number matches the required pattern.
    public bool IsValid => Regex.IsMatch(PhoneNumber, @"^\d{3}-\d{3}-\d{4}$");
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new PhoneNumberModel();
        model.Persons.Add(new Person { PhoneNumber = "123-456-7890" }); // valid
        model.Persons.Add(new Person { PhoneNumber = "5551234" });      // invalid
        model.Persons.Add(new Person { PhoneNumber = "987-654-3210" }); // valid

        // Create a template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach loop over the collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Phone: <<[person.PhoneNumber]>> ");

        // Render "Valid" if the phone number matches the pattern.
        builder.Writeln("<<if [person.IsValid]>>Valid<</if>>");

        // Render "Invalid" if the phone number does not match the pattern.
        builder.Writeln("<<if [!person.IsValid]>>Invalid<</if>>");

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template (optional, shown for completeness).
        const string templatePath = "PhoneNumberTemplate.docx";
        doc.Save(templatePath);

        // Load the template and build the report.
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the final report.
        const string outputPath = "PhoneNumberReport.docx";
        loadedDoc.Save(outputPath);
    }
}
