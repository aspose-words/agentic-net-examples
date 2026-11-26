using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // No properties needed for this example.
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a Word document template with a footer containing
        // the page number field {=PageNumber}.
        // -----------------------------------------------------------------
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Move the builder to the primary footer of the first (and only) section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Write the footer text and insert the page number field.
        builder.Write("Page ");
        // Insert a formula field that evaluates to the current page number.
        builder.InsertField("=PageNumber", null);
        builder.Writeln();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new(templatePath);

        // Create a simple data model (empty in this case).
        ReportModel model = new();

        // Build the report. The root object name in the template is "model".
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Step 3: Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
