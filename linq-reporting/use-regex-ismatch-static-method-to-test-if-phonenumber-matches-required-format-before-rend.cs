using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class PhoneReportModel
{
    public string PhoneNumber { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new PhoneReportModel
        {
            PhoneNumber = "123-456-7890" // Change to test different formats.
        };

        // Create a template document programmatically.
        string templatePath = "PhoneTemplate.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert placeholders and conditional tags.
        builder.Writeln("Phone: <<[model.PhoneNumber]>>");
        builder.Writeln("<<if [Regex.IsMatch(model.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>Valid<</if>>");
        builder.Writeln("<<if [!Regex.IsMatch(model.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>Invalid<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        engine.KnownTypes.Add(typeof(Regex)); // Allow use of Regex static methods in expressions.

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "PhoneReport.docx";
        doc.Save(outputPath);
    }
}
