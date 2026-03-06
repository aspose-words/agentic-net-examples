using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to insert an HTML heading into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            // The HTML string contains an <h1> tag which will be rendered as a heading.
            builder.InsertHtml("<h1>LINQ Reporting Introduction to LINQ Reporting Engine</h1>");

            // Prepare a dummy data source. ReportingEngine requires a data source object,
            // even if the template does not contain any data placeholders.
            var dummyDataSource = new { };

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. This step processes any LINQ Reporting tags in the template.
            // Since the template only contains static HTML, the method simply returns true.
            engine.BuildReport(doc, dummyDataSource);

            // Save the resulting document to disk.
            doc.Save("LINQReportingHeading.docx");
        }
    }
}
