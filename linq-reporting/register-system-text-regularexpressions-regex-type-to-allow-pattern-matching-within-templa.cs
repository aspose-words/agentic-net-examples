using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model used by the template.
    public class Model
    {
        public string Email { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Paths for the temporary template and the final report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write the email field.
        builder.Writeln("Email: <<[model.Email]>>");

        // Validation using System.Text.RegularExpressions.Regex.
        // The regular expression checks for a very simple e‑mail pattern.
        // If the pattern matches, the text is shown in green, otherwise in red.
        // Note: In the template the backslash must be escaped as "\\" for the engine,
        // therefore we use four backslashes in the C# string literal.
        builder.Writeln(
            "<<if [Regex.IsMatch(model.Email, \"^[^@]+@[^@]+\\\\.[^@]+$\")]>>" +
            "<<textColor [\"Green\"]>>Valid Email<</textColor>><</if>>");

        builder.Writeln(
            "<<if [!Regex.IsMatch(model.Email, \"^[^@]+@[^@]+\\\\.[^@]+$\")]>>" +
            "<<textColor [\"Red\"]>>Invalid Email<</textColor>><</if>>");

        // Save the template to disk (required before building the report).
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back and prepare the reporting engine.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Register the Regex type so that its static members can be used in expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // -----------------------------------------------------------------
        // 3. Create sample data.
        // -----------------------------------------------------------------
        Model data = new Model
        {
            Email = "john.doe@example.com" // Change this value to test validation.
        };

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        engine.BuildReport(doc, data, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(reportPath);
    }
}
