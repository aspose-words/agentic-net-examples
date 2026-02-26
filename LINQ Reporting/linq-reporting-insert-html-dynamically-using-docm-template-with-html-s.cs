using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsHtmlInsertDemo
{
    // Simple data source class that contains an HTML string.
    public class HtmlDataSource
    {
        // The property name will be used in the template.
        public string HtmlContent { get; set; }

        public HtmlDataSource(string html)
        {
            HtmlContent = html;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCM template that contains a reporting tag with the html switch,
            // e.g. <<[ds.HtmlContent]:html>>
            const string templatePath = @"C:\Templates\ReportTemplate.docm";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare the HTML that we want to insert dynamically.
            string htmlFragment = @"
                <h2 style='color:#2E86C1;'>Sales Summary</h2>
                <p>The total sales for <b>Q1 2026</b> are <span style='color:#28B463;'>$1,245,300</span>.</p>
                <ul>
                    <li>North America: $560,000</li>
                    <li>Europe: $420,000</li>
                    <li>Asia: $265,300</li>
                </ul>";

            // Create the data source instance.
            HtmlDataSource data = new HtmlDataSource(htmlFragment);

            // Initialise the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument is the name used in the template to reference the data source.
            // In the template you would use <<[ds.HtmlContent]:html>> where "ds" matches the name below.
            engine.BuildReport(template, data, "ds");

            // Save the generated document.
            const string outputPath = @"C:\Output\GeneratedReport.docx";
            template.Save(outputPath);
        }
    }
}
