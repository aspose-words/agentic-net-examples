using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model for the report.
    public class ReportItem
    {
        public string Title { get; set; }
        // HTML fragment that will be inserted into the document.
        public string HtmlContent { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the WORDML (XML) template. The template should contain:
            // 1. A bookmark named "HtmlPlaceholder" where the HTML will be inserted.
            // 2. A LINQ Reporting placeholder for the title, e.g. <<[item.Title]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.xml";
            const string outputPath   = @"C:\Output\GeneratedReport.docx";

            // Load the WORDML template.
            Document doc = new Document(templatePath);

            // Prepare the data source. In a real scenario this could come from a database or LINQ query.
            var data = new List<ReportItem>
            {
                new ReportItem
                {
                    Title = "Quarterly Summary",
                    HtmlContent = @"
                        <h2 style='color:#2E86C1;'>Revenue Overview</h2>
                        <p>The <b>total revenue</b> for this quarter increased by <span style='color:green;'>15%</span> compared to the previous quarter.</p>
                        <ul>
                            <li>Product A: $1,200,000</li>
                            <li>Product B: $850,000</li>
                            <li>Product C: $430,000</li>
                        </ul>"
                }
            };

            // Build the report using LINQ Reporting Engine.
            // The data source name "item" is used inside the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "item");

            // After the report is built, insert the HTML fragment at the bookmark.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToBookmark("HtmlPlaceholder");
            // InsertHtml parses the HTML and converts it to Word formatting.
            builder.InsertHtml(data[0].HtmlContent);

            // Save the final document.
            doc.Save(outputPath);
        }
    }
}
