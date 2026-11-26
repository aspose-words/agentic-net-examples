using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string PhoneNumber { get; set; } = "";
    public bool IsPhoneValid { get; set; }

    public ReportModel(string phoneNumber)
    {
        PhoneNumber = phoneNumber;
        // Simple US phone number validation (e.g., 123-456-7890 or (123) 456-7890)
        IsPhoneValid = Regex.IsMatch(phoneNumber, @"^\(?\d{3}\)?[-\s]?\d{3}[-]\d{4}$");
    }
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Phone Number Report");
        builder.Writeln("<<if [model.IsPhoneValid]>>");
        builder.Writeln("Valid Phone: <<[model.PhoneNumber]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<<if [!model.IsPhoneValid]>>");
        builder.Writeln("Invalid Phone Number");
        builder.Writeln("<</if>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation
        // -------------------------------------------------
        Document doc = new Document(templatePath);

        // Sample data with a phone number to test
        ReportModel model = new ReportModel("123-456-7890"); // Change to an invalid format to test

        // Build the report using LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        doc.Save(outputPath);
    }
}
