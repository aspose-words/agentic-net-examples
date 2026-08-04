using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class PhoneModel
{
    public string PhoneNumber { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a line that conditionally shows whether the phone number matches the pattern.
        // The pattern expects format: 123-456-7890
        builder.Writeln("Phone: <<if [Regex.IsMatch(model.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>");
        builder.Writeln("<<[model.PhoneNumber]>> (valid)");
        builder.Writeln("<</if>>");
        builder.Writeln("<<if [!Regex.IsMatch(model.PhoneNumber, \"^\\\\d{3}-\\\\d{3}-\\\\d{4}$\")]>>");
        builder.Writeln("<<[model.PhoneNumber]>> (invalid)");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data source.
        // -----------------------------------------------------------------
        PhoneModel model = new PhoneModel
        {
            // Change this value to test different formats.
            PhoneNumber = "123-456-7890"
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // Explicitly set options (none in this case).
            Options = ReportBuildOptions.None
        };

        // Allow the engine to use static members of Regex in expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report. The root object name used in the template is "model".
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
