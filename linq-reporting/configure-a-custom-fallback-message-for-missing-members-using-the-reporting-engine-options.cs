using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model – does NOT contain a 'customer' property.
    public class ReportWrapper
    {
        public string Title { get; set; } = "Sample Report";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // This tag references a missing member (customer.Name). It will be replaced
            // by the custom fallback message configured later.
            builder.Writeln("Customer: <<[customer.Name]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Allow missing members to be treated as null and use the custom message.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = "Data not available";

            // Build the report using a data source that does NOT contain 'customer'.
            // The overload without a data source name allows direct member access.
            engine.BuildReport(reportDoc, new ReportWrapper());

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);

            // Inform the user where the files are located.
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report generated at: {outputPath}");
        }
    }
}
