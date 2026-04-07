using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var person = new Person
        {
            PhoneNumber = "123-456-7890" // Change to an invalid format to see the alternative output.
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Add a simple heading.
        builder.Writeln("Contact Information");
        builder.Writeln();

        // Conditional rendering using Regex.IsMatch to validate the phone number format.
        // The pattern expects the format XXX-XXX-XXXX where X is a digit.
        builder.Writeln(
            "<<if [Regex.IsMatch(PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>" +
            "Phone: <<[PhoneNumber]>>" +
            "<</if>>" +
            "<<if [!Regex.IsMatch(PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>" +
            "Phone: Invalid" +
            "<</if>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for reporting.
        var loadedDoc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register the Regex type so its static members can be used in the template.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report using the root object name "person".
        engine.BuildReport(loadedDoc, person, "person");

        // Save the generated report.
        var outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}

// Simple data model with a phone number property.
public class Person
{
    public string PhoneNumber { get; set; } = string.Empty;
}
