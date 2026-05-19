using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Returns the name of the color to be applied to the text.
        public string ColorName { get; set; } = "Blue";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that changes the text color dynamically.
            // The color expression reads the ColorName property from the model.
            builder.Writeln("<<textColor [model.ColorName]>>Status Text<</textColor>>");

            // Save the template to disk so it can be loaded later.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Prepare the data source.
            ReportModel model = new ReportModel();

            // Create the reporting engine and generate the report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // -------------------------------------------------
            // 3. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
