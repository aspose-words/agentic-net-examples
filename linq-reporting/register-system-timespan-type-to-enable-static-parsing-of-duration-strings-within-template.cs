using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple wrapper class required as a root data source for the report.
    public class DummyRoot
    {
        // No members needed for this example.
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document with a LINQ Reporting tag.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // The tag uses the static TimeSpan.Parse method to convert a duration string.
            // Use double quotes inside the expression to avoid char literal parsing errors.
            builder.Writeln("Parsed duration: <<[TimeSpan.Parse(\"02:15:30\")]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template document for reporting.
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Register System.TimeSpan so that static parsing can be used in the template.
            engine.KnownTypes.Add(typeof(TimeSpan));

            // -------------------------------------------------
            // 4. Build the report.
            // -------------------------------------------------
            // The template does not reference any data source members, but BuildReport still requires one.
            // We pass an instance of DummyRoot with a root name that does not appear in the template.
            engine.BuildReport(reportDoc, new DummyRoot(), "dummy");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(reportPath);
        }
    }
}
