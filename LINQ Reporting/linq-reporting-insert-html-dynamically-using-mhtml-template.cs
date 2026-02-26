using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the template and the output file. Adjust as needed.
            string templatePath = "Template.mhtml";   // MHTML template that contains <<[ds.HtmlContent]>>
            string outputPath   = "Result.mhtml";    // Where the generated MHTML will be saved

            // Dynamic HTML that will be inserted into the template.
            string htmlContent = "<p><b>Hello, Aspose.Words!</b></p>";

            // Create the generator and build the report.
            var generator = new HtmlReportGenerator();
            generator.Generate(templatePath, outputPath, htmlContent);

            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }

    public class HtmlReportGenerator
    {
        /// <summary>
        /// Generates a report by loading an MHTML template, inserting dynamic HTML,
        /// and saving the result as an MHTML document.
        /// </summary>
        /// <param name="templatePath">Full path to the MHTML template file.</param>
        /// <param name="outputPath">Full path where the generated MHTML will be saved.</param>
        /// <param name="htmlContent">HTML string to be inserted into the template.</param>
        public void Generate(string templatePath, string outputPath, string htmlContent)
        {
            // Load the MHTML template into an Aspose.Words Document.
            Document doc = new Document(templatePath);

            // Create a simple anonymous data source containing the HTML string.
            // The template should reference this value, e.g. <<[ds.HtmlContent]>>.
            var dataSource = new { HtmlContent = htmlContent };

            // Build the report using the ReportingEngine.
            // The data source name ("ds") must match the name used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "ds");

            // Save the populated document as MHTML.
            doc.Save(outputPath, SaveFormat.Mhtml);
        }
    }
}
