using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the template document.
            Document template = new Document("Template.docx");

            // Create a reporting engine instance.
            ReportingEngine engine = new ReportingEngine();

            // Enable the option to remove empty paragraphs after the report is built.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Prepare a data source for the report (replace with your actual data source).
            // For demonstration, we use an empty DataSet.
            var dataSource = new System.Data.DataSet();

            // Build the report using the template and the data source.
            engine.BuildReport(template, dataSource);

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
