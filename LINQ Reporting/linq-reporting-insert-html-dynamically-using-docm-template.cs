using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains a merge field <<[ds.HtmlContent]>>.
            // The template can be created in Word and saved as a macro‑enabled document.
            Document doc = new Document(@"C:\Templates\ReportTemplate.docm");

            // Define a simple data source with a title and an HTML fragment.
            var dataSource = new
            {
                Title = "Quarterly Sales Report",
                HtmlContent = "<h2>Summary</h2><p><b>Revenue:</b> $1,200,000</p>"
            };

            // Populate the template using the LINQ Reporting Engine.
            // The data source name "ds" is used inside the template (e.g. <<[ds.Title]>>).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "ds");

            // After the report is built, replace the merge field that holds the HTML
            // with the actual formatted content using DocumentBuilder.InsertHtml.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToMergeField("HtmlContent");          // Move cursor to the field.
            builder.InsertHtml(dataSource.HtmlContent);       // Insert the HTML fragment.

            // Save the final document. The output format is inferred from the file extension.
            doc.Save(@"C:\Output\QuarterlyReport.docx");
        }
    }
}
