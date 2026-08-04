using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // External types whose static members will be accessed from the template.
    public static class ExternalA
    {
        public static string ValueA => "Hello from ExternalA";
    }

    public static class ExternalB
    {
        public static int ValueB => 42;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags that refer
            //    to static members of the external types.
            // -----------------------------------------------------------------
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Static values accessed via KnownTypes:");
            builder.Writeln("ExternalA.ValueA = <<[ExternalA.ValueA]>>");
            builder.Writeln("ExternalB.ValueB = <<[ExternalB.ValueB]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back for report generation.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            //    Register the external types so the template can use them without
            //    reflection.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(ExternalA));
            engine.KnownTypes.Add(typeof(ExternalB));

            // Build the report. No root data object is required because the template
            // only uses static members.
            engine.BuildReport(reportDoc, new object(), "data");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(reportPath);

            // Output the resulting text to the console for verification.
            Console.WriteLine("Report generated successfully. Content:");
            Console.WriteLine(reportDoc.GetText());
        }
    }
}
