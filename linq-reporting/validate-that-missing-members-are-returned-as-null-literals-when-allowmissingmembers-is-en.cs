using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string outputPath = "Output.docx";

            // -------------------------------------------------
            // 1. Create a template document that references a missing member.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);
            builder.Writeln("Missing member test: <<[missingObject.First().id]>>");
            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template back (required by the lifecycle rule).
            // -------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare a data source that does NOT contain the missing member.
            // -------------------------------------------------
            DataSet emptyDataSet = new DataSet(); // No tables, no members.

            // -------------------------------------------------
            // 4. Configure the ReportingEngine to allow missing members.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.AllowMissingMembers,
                MissingMemberMessage = "Missed"
            };

            // Build the report. The root name is empty because we don't reference the root object itself.
            engine.BuildReport(reportDoc, emptyDataSet, "");

            // -------------------------------------------------
            // 5. Save the generated report.
            // -------------------------------------------------
            reportDoc.Save(outputPath);

            // -------------------------------------------------
            // 6. Verify that the missing member was replaced with the custom message.
            // -------------------------------------------------
            string resultText = reportDoc.GetText();
            Console.WriteLine("Report generated. Contains custom missing‑member message? " + resultText.Contains("Missed"));
        }
    }
}
