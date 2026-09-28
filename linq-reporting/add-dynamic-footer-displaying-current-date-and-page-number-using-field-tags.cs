using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Define paths for the template and the final report.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with a footer that contains
        //    DATE, PAGE, and NUMPAGES fields.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Move to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // Insert the desired fields directly (LINQ Reporting tags are not needed for Word fields).
        builder.Write("Date: ");
        builder.InsertField("DATE \\@ \"yyyy-MM-dd\"");
        builder.Write("  Page: ");
        builder.InsertField("PAGE");
        builder.Write(" of ");
        builder.InsertField("NUMPAGES");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // The model is empty because the footer does not depend on external data.
        ReportModel model = new();

        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 3. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}

// Empty model class required by the ReportingEngine.
public class ReportModel
{
    // No properties needed for this example.
}
