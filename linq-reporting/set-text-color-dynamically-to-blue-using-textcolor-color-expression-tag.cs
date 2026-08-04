using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // The color expression used by the <<textColor>> tag.
        public string ColorName { get; set; } = "Blue";

        // Any additional data can be added here.
        public string Message { get; set; } = "Status Text";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and final report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that sets the text color dynamically.
            // The tag uses the model's ColorName property to determine the color.
            builder.Writeln("<<textColor [model.ColorName]>>Status Text<</textColor>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(loadedTemplate, model, "model");

            // Save the generated report.
            loadedTemplate.Save(reportPath);
        }
    }
}
