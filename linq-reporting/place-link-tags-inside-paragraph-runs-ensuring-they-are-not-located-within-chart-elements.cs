using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
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

        // Paragraph that will contain the link tag.
        builder.Writeln("<<link [Url] [Text]>>");

        // Insert a chart – the link tag must NOT be placed inside the chart.
        // The ChartType enum is defined in Aspose.Words.Drawing.Charts.
        builder.InsertChart(ChartType.Column, 400, 300);
        // (No LINQ Reporting tags are added inside the chart.)

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template back for reporting.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Url = "https://example.com",
            Text = "Example Site"
        };

        // -----------------------------------------------------------------
        // 4. Build the report using Aspose.Words LINQ Reporting Engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // The template uses <<link [Url] [Text]>>, so we pass the root name "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        loadedTemplate.Save(reportPath);
    }
}

// ---------------------------------------------------------------------
// Data model referenced by the template.
// ---------------------------------------------------------------------
public class ReportModel
{
    // URL for the hyperlink.
    public string Url { get; set; } = string.Empty;

    // Display text for the hyperlink.
    public string Text { get; set; } = string.Empty;
}
