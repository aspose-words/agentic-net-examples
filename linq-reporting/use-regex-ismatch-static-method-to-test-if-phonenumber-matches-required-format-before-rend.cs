using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample phone number property.
    public string PhoneNumber { get; set; } = string.Empty;

    // Computed property that validates the phone number format.
    public bool IsPhoneValid => Regex.IsMatch(PhoneNumber, @"^\d{3}-\d{3}-\d{4}$");
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Write a line that conditionally shows "Valid" or "Invalid"
        // based on whether PhoneNumber matches the required pattern.
        builder.Writeln(
            "Phone: <<if [IsPhoneValid]>>Valid<</if>><<if [!IsPhoneValid]>>Invalid<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Sample data model with a phone number.
        var model = new ReportModel
        {
            PhoneNumber = "123-456-7890" // Change to test different formats.
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
