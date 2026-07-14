using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Model class used as the root data source for the LINQ Reporting template.
    public class ReportModel
    {
        // Holds the current UTC time at the moment the model is instantiated.
        public DateTime CurrentUtc { get; set; } = DateTime.UtcNow;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            const string templatePath = "template.docx";
            const string reportPath = "report.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a simple line that will display the current UTC time.
            // The tag <<[model.CurrentUtc]>> references the property of the root object.
            builder.Writeln("Current UTC time: <<[model.CurrentUtc]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Create the data model. The property is already populated with UtcNow.
            ReportModel model = new ReportModel();

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Allow the engine to resolve static members of DateTime if needed.
            engine.KnownTypes.Add(typeof(DateTime));

            // Build the report. The root name "model" must match the tag used in the template.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 3. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
