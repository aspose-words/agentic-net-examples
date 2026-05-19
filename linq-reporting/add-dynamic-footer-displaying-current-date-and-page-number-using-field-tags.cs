using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize with the current date to avoid nullable warnings.
    public string CurrentDate { get; set; } = DateTime.Now.ToString("d");
}

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string outputPath   = Path.Combine(Environment.CurrentDirectory, "ReportWithFooter.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document that contains a footer with LINQ tags.
        // -----------------------------------------------------------------
        var templateDoc   = new Document();
        var builder       = new DocumentBuilder(templateDoc);

        // Move the cursor to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the date expression tag (bound to the model) and a Word PAGE field.
        builder.Write("Date: <<[model.CurrentDate]>>   Page: ");
        builder.InsertField("PAGE"); // Standard Word field for page number.

        // Save the template so it can be loaded later.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and generate the report using the LINQ engine.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Prepare the data model.
        var model = new ReportModel();

        // Build the report. The root name used in the template tags is "model".
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save(outputPath);
    }
}
