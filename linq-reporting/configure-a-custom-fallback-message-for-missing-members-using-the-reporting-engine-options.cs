using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class without the expected member.
    public class EmptyData
    {
        // No properties – the template will reference a missing member.
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Insert a LINQ Reporting tag that references a non‑existent member.
            // The tag uses the root name "data".
            builder.Writeln("Hello <<[data.MissingMember]>>!");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back from disk (required by the workflow).
            // -----------------------------------------------------------------
            Document loadedTemplate = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();

            // Allow missing members and provide a custom fallback message.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            engine.MissingMemberMessage = "[Member not found]";

            // Build the report using an instance of EmptyData as the data source.
            // The root name must match the one used in the template ("data").
            engine.BuildReport(loadedTemplate, new EmptyData(), "data");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            loadedTemplate.Save(reportPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
        }
    }
}
