using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public string PhoneNumber { get; set; } = "";

    // Determines whether the phone number matches the required format.
    public bool IsPhoneValid => Regex.IsMatch(PhoneNumber, @"^\d{3}-\d{3}-\d{4}$");
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for additional encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // LINQ Reporting tags.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("<<if [person.IsPhoneValid]>>Phone: <<[person.PhoneNumber]>> (valid)<</if>>");
        builder.Writeln("<<if [!person.IsPhoneValid]>>Phone: <<[person.PhoneNumber]>> (invalid)<</if>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data.
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "Alice", PhoneNumber = "123-456-7890" }, // valid
                new Person { Name = "Bob",   PhoneNumber = "9876543210" }   // invalid
            }
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
