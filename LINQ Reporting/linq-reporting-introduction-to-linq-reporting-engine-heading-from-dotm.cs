using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the DOTM template that contains the heading.
            // The template should have the text "LINQ Reporting Introduction to LINQ Reporting Engine"
            // possibly wrapped in LINQ Reporting Engine tags.
            string templatePath = @"C:\Templates\LinqReportingTemplate.dotm";

            // Load the DOTM template document.
            Document doc = new Document(templatePath);

            // Create an instance of the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. No external data source is required for a static heading,
            // so we pass an empty anonymous object as the data source.
            // The third parameter (dataSourceName) can be null because we do not reference the
            // data source object itself in the template.
            engine.BuildReport(doc, new { }, null);

            // Save the populated document as a DOCX file.
            string outputPath = @"C:\Output\LinqReportingResult.docx";
            doc.Save(outputPath);
        }
    }
}
