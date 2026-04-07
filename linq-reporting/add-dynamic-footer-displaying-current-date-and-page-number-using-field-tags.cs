using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // Current date is initialized when the object is created.
        public string CurrentDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
    }

    public static void Main()
    {
        // Paths for the temporary template and the final report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Add some body content so the document has at least one page.
        builder.Writeln("This is a sample report generated with Aspose.Words LINQ Reporting.");
        builder.Writeln("The footer below shows the current date and the page number.");

        // -----------------------------------------------------------------
        // 2. Create a footer that contains a LINQ Reporting tag for the date
        //    and a Word field for the page number.
        // -----------------------------------------------------------------
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

        // LINQ Reporting tag – will be replaced with the value of ReportModel.CurrentDate.
        builder.Write("Date: <<[CurrentDate]>>  ");

        // Word field – displays the current page number and updates automatically.
        builder.InsertField("PAGE \\* MERGEFORMAT");

        // Return the cursor to the main document body.
        builder.MoveToDocumentEnd();

        // Save the template to disk (required before building the report).
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the template and build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data source.
        ReportModel model = new ReportModel();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            // No special options are needed for this simple scenario.
            Options = ReportBuildOptions.None
        };

        // Build the report. The root object name ("model") must match the tag usage.
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}
