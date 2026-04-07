using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingIncludeExample
{
    // Simple data model used by the template.
    public class ReportModel
    {
        public string Title { get; set; } = "Quarterly Sales Report";
        public string Date { get; set; } = DateTime.Today.ToString("D");
        public string Author { get; set; } = "John Doe";
        public string Body { get; set; } = "This is the body of the report generated using Aspose.Words LINQ Reporting.";
    }

    public class Program
    {
        // File names used in the example.
        private const string HeaderFragmentFile = "HeaderFragment.docx";
        private const string MainTemplateFile = "MainTemplate.docx";
        private const string OutputReportFile = "ReportOutput.docx";

        public static void Main()
        {
            // Required for some encodings used by Aspose.Words.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // 1. Create a reusable header fragment.
            CreateHeaderFragment();

            // 2. Create the main template that will contain the header fragment.
            CreateMainTemplate();

            // 3. Load the main template.
            Document template = new Document(MainTemplateFile);

            // 4. Prepare the data model.
            ReportModel model = new ReportModel();

            // 5. Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None // default options
            };
            engine.BuildReport(template, model, "model");

            // 6. Save the generated report.
            template.Save(OutputReportFile);
        }

        // Creates a separate document that contains common header information.
        private static void CreateHeaderFragment()
        {
            Document headerDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(headerDoc);

            // Header fields that will be filled from the data model.
            builder.Writeln("Report Title: <<[model.Title]>>");
            builder.Writeln("Date: <<[model.Date]>>");
            builder.Writeln("Author: <<[model.Author]>>");

            // Save the fragment to disk.
            headerDoc.Save(HeaderFragmentFile);
        }

        // Creates the main template and inserts the header fragment programmatically.
        private static void CreateMainTemplate()
        {
            Document mainDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(mainDoc);

            // Load the previously created header fragment.
            Document headerFragment = new Document(HeaderFragmentFile);

            // Insert the header fragment into the main template.
            builder.InsertDocument(headerFragment, ImportFormatMode.KeepSourceFormatting);

            // Add a blank line and body content placeholder.
            builder.Writeln();
            builder.Writeln("<<[model.Body]>>");

            // Save the main template.
            mainDoc.Save(MainTemplateFile);
        }
    }
}
