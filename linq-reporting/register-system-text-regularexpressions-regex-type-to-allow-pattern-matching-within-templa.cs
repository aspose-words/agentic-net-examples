using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model with an Email property.
    public class Model
    {
        public string Email { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a placeholder for the email value.
        builder.Writeln("Email: <<[model.Email]>>");

        // Write a validation message that appears only when the email matches the regex pattern.
        // The pattern checks for a simple email format.
        // Note: Backslashes in the regex need to be escaped twice:
        //   - once for the C# string literal,
        //   - once for the LINQ Reporting engine parser.
        builder.Writeln(
            "<<if [Regex.IsMatch(model.Email, \"^[^@\\\\\\\\s]+@[^@\\\\\\\\s]+\\\\\\\\.[^@\\\\\\\\s]+$\")]>>Valid Email<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Prepare the data source.
        Model data = new Model { Email = "john.doe@example.com" }; // Change to an invalid email to see the message omitted.

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Register the Regex type so that template expressions can use its static methods.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report using the model object. The root name in the template is "model".
        engine.BuildReport(reportDoc, data, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
